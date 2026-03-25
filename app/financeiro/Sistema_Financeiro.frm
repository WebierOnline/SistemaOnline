VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Sistema_Financeiro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONFIGURAÇÕES"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   Icon            =   "Sistema_Financeiro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9915
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   17489
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
      TabPicture(0)   =   "Sistema_Financeiro.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMarcado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStatus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdImprimir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGerenciaNet"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdMarcarTodos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEnviarTodos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ccmdMostrarRazao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdMostrarTudo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdEnviarUm"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdDesmarcarTodos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdDesmarcar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdMarcar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Grid"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frmCadastro"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "0"
      TabPicture(1)   =   "Sistema_Financeiro.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "0"
      TabPicture(2)   =   "Sistema_Financeiro.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSalvarBalanca"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "0"
      TabPicture(3)   =   "Sistema_Financeiro.frx":2426
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "0"
      TabPicture(4)   =   "Sistema_Financeiro.frx":2442
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Mensagens para clientes"
         Height          =   1875
         Left            =   60
         TabIndex        =   33
         Top             =   7920
         Width           =   11955
         Begin VB.TextBox txtMensagem 
            Height          =   675
            Left            =   60
            TabIndex        =   37
            Text            =   "Bom dia {cliente} Segue anexo a msg acima citadas {codigodesbloqueio}"
            Top             =   480
            Width           =   6675
         End
         Begin VB.TextBox txtCaminho 
            Height          =   285
            Left            =   60
            MaxLength       =   200
            TabIndex        =   35
            Top             =   1500
            Width           =   5715
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   11400
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizarArquivo 
            Height          =   315
            Left            =   5820
            TabIndex        =   42
            Top             =   1500
            Width           =   915
            _ExtentX        =   1614
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Sistema_Financeiro.frx":245E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarMsgUm 
            Height          =   315
            Left            =   7440
            TabIndex        =   44
            Top             =   420
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar Msg Selec"
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
            MICON           =   "Sistema_Financeiro.frx":247A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarMsgTodos 
            Height          =   315
            Left            =   7440
            TabIndex        =   47
            Top             =   780
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar Msg Todos"
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
            MICON           =   "Sistema_Financeiro.frx":2496
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarAnexoUm 
            Height          =   315
            Left            =   7440
            TabIndex        =   49
            Top             =   1140
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar Anexo Selec"
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
            MICON           =   "Sistema_Financeiro.frx":24B2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarMsgAnexoUm 
            Height          =   315
            Left            =   9300
            TabIndex        =   50
            Top             =   420
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar Msg e Anexo Selec"
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
            MICON           =   "Sistema_Financeiro.frx":24CE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarMsgAnexoTodos 
            Height          =   315
            Left            =   9300
            TabIndex        =   51
            Top             =   780
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar Msg e Anexo Todos"
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
            MICON           =   "Sistema_Financeiro.frx":24EA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdLerQRCode 
            Height          =   315
            Left            =   9300
            TabIndex        =   64
            Top             =   1140
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Ler QRCode"
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
            MICON           =   "Sistema_Financeiro.frx":2506
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarTodosUm 
            Height          =   315
            Left            =   7440
            TabIndex        =   65
            Top             =   1500
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar Anexo Todos"
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
            MICON           =   "Sistema_Financeiro.frx":2522
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdDesconectarWhats 
            Height          =   315
            Left            =   9300
            TabIndex        =   66
            Top             =   1500
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Desconectar"
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
            MICON           =   "Sistema_Financeiro.frx":253E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Shape imgOFFLine 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000FF&
            FillColor       =   &H000000FF&
            Height          =   375
            Left            =   11400
            Shape           =   3  'Circle
            Top             =   1380
            Width           =   495
         End
         Begin VB.Shape imgOnLine 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H0000FF00&
            FillColor       =   &H0000FF00&
            Height          =   375
            Left            =   11400
            Shape           =   3  'Circle
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anexo:"
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
            TabIndex        =   53
            Top             =   1260
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mensagem"
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
            TabIndex        =   52
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Desbloqueio Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   60
         TabIndex        =   15
         Top             =   7020
         Width           =   11955
         Begin VB.TextBox txtCodDesbloqueioTemp 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtCodDesbloqueio 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   4020
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1680
            TabIndex        =   17
            Top             =   480
            Width           =   1155
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1515
         End
         Begin ChamaleonBtn.chameleonButton cmdMostrarSenha 
            Height          =   315
            Left            =   2880
            TabIndex        =   18
            Top             =   480
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
            MICON           =   "Sistema_Financeiro.frx":255A
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
            TabIndex        =   20
            Top             =   480
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
            MICON           =   "Sistema_Financeiro.frx":2576
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
            TabIndex        =   22
            Top             =   480
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
            MICON           =   "Sistema_Financeiro.frx":2592
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdGerarMes 
            Height          =   315
            Left            =   6840
            TabIndex        =   23
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Gerar Mês Atual"
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
            MICON           =   "Sistema_Financeiro.frx":25AE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdGerarCodigos 
            Height          =   315
            Left            =   8460
            TabIndex        =   24
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Gerar Códigos"
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
            MICON           =   "Sistema_Financeiro.frx":25CA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label60 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Temporário"
            Height          =   195
            Left            =   5460
            TabIndex        =   31
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Certo"
            Height          =   195
            Left            =   4020
            TabIndex        =   29
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mês"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   300
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ano"
            Height          =   195
            Left            =   1680
            TabIndex        =   25
            Top             =   240
            Width           =   285
         End
      End
      Begin VB.Frame frmCadastro 
         Caption         =   "Cadastro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   60
         TabIndex        =   12
         Top             =   4980
         Width           =   11955
         Begin VB.OptionButton optCPF 
            Caption         =   "CPF"
            Height          =   195
            Left            =   9780
            TabIndex        =   63
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optCNPJ 
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   9000
            TabIndex        =   62
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.ComboBox cboCidade 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2100
            TabIndex        =   40
            Top             =   1080
            Width           =   1515
         End
         Begin VB.ComboBox cboEstado 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3660
            TabIndex        =   41
            Top             =   1080
            Width           =   1155
         End
         Begin VB.TextBox txtFantasia 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   480
            Width           =   3435
         End
         Begin VB.TextBox txtRazao 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3960
            TabIndex        =   30
            Top             =   480
            Width           =   4575
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   315
            Left            =   11460
            TabIndex        =   38
            Top             =   480
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
            MICON           =   "Sistema_Financeiro.frx":25E6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskCPF 
            Height          =   315
            Left            =   9000
            TabIndex        =   34
            Top             =   480
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   315
            Left            =   3540
            TabIndex        =   28
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "L"
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
            MICON           =   "Sistema_Financeiro.frx":2602
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton2 
            Height          =   315
            Left            =   8580
            TabIndex        =   32
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "L"
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
            MICON           =   "Sistema_Financeiro.frx":261E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton3 
            Height          =   315
            Left            =   11040
            TabIndex        =   36
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "L"
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
            MICON           =   "Sistema_Financeiro.frx":263A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionar 
            Height          =   315
            Left            =   8580
            TabIndex        =   45
            Top             =   1080
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
            MICON           =   "Sistema_Financeiro.frx":2656
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
            Left            =   7500
            TabIndex        =   43
            Top             =   1080
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
            MICON           =   "Sistema_Financeiro.frx":2672
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
            Height          =   315
            Left            =   10740
            TabIndex        =   48
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
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
            MICON           =   "Sistema_Financeiro.frx":268E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdEditar 
            Height          =   315
            Left            =   9600
            TabIndex        =   46
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Editar"
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
            MICON           =   "Sistema_Financeiro.frx":26AA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskCelular 
            Height          =   315
            Left            =   60
            TabIndex        =   39
            Top             =   1080
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton4 
            Height          =   315
            Left            =   4860
            TabIndex        =   60
            Top             =   1080
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "L"
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
            MICON           =   "Sistema_Financeiro.frx":26C6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   3660
            TabIndex        =   59
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   2100
            TabIndex        =   58
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Celular"
            Height          =   195
            Left            =   60
            TabIndex        =   54
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fantasia"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Razão"
            Height          =   195
            Left            =   4020
            TabIndex        =   13
            Top             =   240
            Width           =   465
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarBalanca 
         Height          =   615
         Left            =   -68400
         TabIndex        =   5
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
         MICON           =   "Sistema_Financeiro.frx":26E2
         PICN            =   "Sistema_Financeiro.frx":26FE
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
         Height          =   3675
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   6482
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdMarcar 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   4380
         Width           =   795
         _ExtentX        =   1402
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
         MICON           =   "Sistema_Financeiro.frx":4490
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
         Left            =   960
         TabIndex        =   1
         Top             =   4380
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "Sistema_Financeiro.frx":44AC
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
         Left            =   3180
         TabIndex        =   2
         Top             =   4380
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
         MICON           =   "Sistema_Financeiro.frx":44C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEnviarUm 
         Height          =   315
         Left            =   8520
         TabIndex        =   7
         Top             =   4380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Enviar Selecionado"
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
         MICON           =   "Sistema_Financeiro.frx":44E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMostrarTudo 
         Height          =   315
         Left            =   9240
         TabIndex        =   8
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Mostrar Fantasia"
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
         MICON           =   "Sistema_Financeiro.frx":4500
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton ccmdMostrarRazao 
         Height          =   315
         Left            =   10680
         TabIndex        =   9
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Mostrar Razao"
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
         MICON           =   "Sistema_Financeiro.frx":451C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEnviarTodos 
         Height          =   315
         Left            =   10140
         TabIndex        =   10
         Top             =   4380
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Enviar Todos Marcados"
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
         MICON           =   "Sistema_Financeiro.frx":4538
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdMarcarTodos 
         Height          =   315
         Left            =   1860
         TabIndex        =   11
         Top             =   4380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Marcar Todos"
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
         MICON           =   "Sistema_Financeiro.frx":4554
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdGerenciaNet 
         Height          =   315
         Left            =   7440
         TabIndex        =   56
         Top             =   4380
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Ler Arquivo"
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
         MICON           =   "Sistema_Financeiro.frx":4570
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
         Left            =   4680
         TabIndex        =   61
         Top             =   4380
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Imprimir"
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
         MICON           =   "Sistema_Financeiro.frx":458C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   4740
         Width           =   750
      End
      Begin VB.Label lblMarcado 
         AutoSize        =   -1  'True
         Caption         =   "0000"
         Height          =   195
         Left            =   5520
         TabIndex        =   55
         Top             =   4440
         Width           =   360
      End
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Confirmar fechamendo da venda:"
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   5040
      Width           =   2355
   End
End
Attribute VB_Name = "Sistema_Financeiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moCombo As cComboHelper
Private Caminho As String
Dim oCfg As ConfigItem
Dim sSQL As String
Dim r As ADODB.Recordset
Dim i As Integer
Dim vMesRef As String
Dim vAnoRef As String

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Sub cboCidade_GotFocus()
    cboCidade.Clear
    
    sSQL = "SELECT cidade FROM empresas_desbloueio GROUP BY cidade;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboCidade.AddItem ValidateNull(r("cidade"))
       r.MoveNext
    Loop
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    moCombo.AttachTo cboCidade
End Sub

Private Sub cboEstado_GotFocus()
    cboEstado.Clear
    
    sSQL = "SELECT Estado FROM empresas_desbloueio GROUP BY estado;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboEstado.AddItem ValidateNull(r("Estado"))
       r.MoveNext
    Loop
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    moCombo.AttachTo cboEstado
End Sub


Private Sub ccmdMostrarRazao_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado, (CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado FROM  empresas_desbloueio ORDER BY RAZAO;"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    sSQL = "SELECT * FROM  empresas_desbloueio where marcado = 1;"
    Set r = dbData.OpenRecordset(sSQL)
    
    lblMarcado.Caption = r.RecordCount
End Sub

Private Sub chameleonButton2_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado,(CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF  FROM  empresas_desbloueio where (RAZAO LIKE '%" & txtRazao.Text & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub chameleonButton3_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado, (CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF FROM  empresas_desbloueio where (CNPJ LIKE '%" & mskCPF.Text & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub chameleonButton4_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado,(CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF  FROM  empresas_desbloueio where (cidade LIKE '%" & cboCidade.Text & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub cmdLerQRCode_Click()
Dim iRetorno As Boolean
    cmdLerQRCode.Enabled = False
    DoEvents
    iRetorno = WhatsAppConectar
    cmdLerQRCode.Enabled = True
    DoEvents
End Sub

Private Sub cmdDesconectarWhats_Click()
Dim iRetorno As Boolean
    cmdDesconectarWhats.Enabled = False
    DoEvents
    iRetorno = WhatsAppDesconectar
    cmdDesconectarWhats.Enabled = True
    DoEvents
End Sub

Private Sub cmdEditar_Click()
    If txtFantasia.Text = "" Or txtRazao.Text = "" Or mskCPF.Text = "" Then Exit Sub
    
    If Not Atualizar_Dados Then
       ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
       Exit Sub
    End If
    
    'MostrarEmpresa
    LimparEmpresa
End Sub

Private Function Atualizar_Dados() As Boolean
i = Grid.Row

sSQL = "UPDATE empresas_desbloueio SET razao = '" & txtRazao.Text & "', fantasia = '" & txtFantasia.Text & "', cnpj = '" & mskCPF.Text & "', celular = '" & mskCelular.Text & "', cidade = '" & cboCidade.Text & "', estado = '" & cboEstado.Text & "', cpf = '" & Abs(optCPF.Value) & "' WHERE (codigo = " & Grid.TextMatrix(i, 4) & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdEnviarAnexoUm_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT RAZAO, CNPJ, celular FROM empresas_desbloueio WHERE Marcado = 1"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviarAnexo(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular)
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0 WHERE Marcado = 1"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarMsgAnexoTodos_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT RAZAO, CNPJ, celular FROM empresas_desbloueio"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviarAnexo(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular)
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0 WHERE Marcado = 1"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarMsgAnexoUm_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT RAZAO, CNPJ, celular FROM empresas_desbloueio WHERE Marcado = 1"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviarAnexo(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular)
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0 WHERE Marcado = 1"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarMsgTodos_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT RAZAO, CNPJ, celular, Cod_Desbloqueio FROM empresas_desbloueio"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviar(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular, "")
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0 WHERE Marcado = 1"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarMsgUm_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT RAZAO, CNPJ, celular, ISNULL(Cod_Desbloqueio, '') AS Cod_Desbloqueio FROM empresas_desbloueio WHERE Marcado = 1"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviar(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular, "")
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0 WHERE Marcado = 1"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarTodos_Click()
Dim rsClientes As ADODB.Recordset
'Dim sSQL As String
    
    sSQL = "SELECT RAZAO, FANTASIA, CNPJ, celular, ISNULL(Cod_Desbloqueio, '') AS Cod_Desbloqueio,  Mes_Referente, Ano_Referente FROM empresas_desbloueio WHERE Marcado = 1"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
        vMesRef = rsClientes!Mes_Referente
        vAnoRef = rsClientes!Ano_Referente
       iRetorno = WhatsAppEnviarCodigo(rsClientes!CNPJ, rsClientes!Fantasia, rsClientes!Celular, rsClientes!Cod_Desbloqueio)
       rsClientes.MoveNext
    Loop
    'sSQL = "UPDATE empresas_desbloueio SET Marcado = 0 WHERE Marcado = 1"
    'SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarTodosUm_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT RAZAO, CNPJ, celular FROM empresas_desbloueio"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviarAnexo(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular)
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0, Enviado = 1 WHERE Marcado = 1"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing
End Sub

Private Sub cmdEnviarUm_Click()
Dim rsClientes As ADODB.Recordset
Dim sSQL As String, cCNPJ As String

    i = Grid.Row

    If Grid.TextMatrix(i, 6) = "" Then
       MsgBox "Nenhum número de whatsapp foi localizado!", vbExclamation, "Enviando Mensagem WhatsApp"
       Exit Sub
    End If

    If Len(Grid.TextMatrix(i, 6)) < 15 Then
       MsgBox "Número de telefone inválido!", vbCritical, "Enviando Mensagem WhatsApp"
       Exit Sub
    End If
    
    cCNPJ = Grid.TextMatrix(i, 3)

    sSQL = "SELECT RAZAO, CNPJ, celular, ISNULL(Cod_Desbloqueio, '') AS Cod_Desbloqueio FROM empresas_desbloueio WHERE CNPJ = '" & cCNPJ & "'"
    RsOpen rsClientes, sSQL
    Do While Not rsClientes.EOF
       iRetorno = WhatsAppEnviarCodigo(rsClientes!CNPJ, rsClientes!Razao, rsClientes!Celular, rsClientes!Cod_Desbloqueio)
       rsClientes.MoveNext
    Loop
    sSQL = "UPDATE empresas_desbloueio SET Marcado = 0, Enviado = 1 WHERE CNPJ = '" & cCNPJ & "'"
    SQLExecuta sSQL
    MostrarEmpresa
    Set rsClientes = Nothing

'Dim vNumCel As String
'vNumCel = Grid.TextMatrix(i, 6)

'vNumCel = Replace(Replace(Replace(vNumCel, "(", ""), ")", ""), "-", "")
'Chama a função ShellExecute = url da api do whatsapp web
'ShellExecute hwnd, "open", ("https://api.whatsapp.com/send?phone=55" & vNumCel), _
'vbNullString, vbNullString, conSwNo

End Sub

Private Sub cmdExcluir_Click()
i = Grid.Row
dbData.Execute "DELETE FROM empresas_desbloueio WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
MostrarEmpresa
End Sub

Private Sub cmdGerarCodigos_Click()
Dim rsClientesMarcados As ADODB.Recordset
    
    On Error GoTo deuErro
    
    'Gera código Desbloqueio
    sSQL = "SELECT RAZAO, CNPJ, celular, Mes_Referente, Ano_Referente, Cod_Desbloqueio FROM empresas_desbloueio WHERE marcado = 1"
    RsOpen rsClientesMarcados, sSQL
    If rsClientesMarcados.RecordCount > 0 Then rsClientesMarcados.MoveLast: rsClientesMarcados.MoveFirst
    Do While Not rsClientesMarcados.EOF
        codDesbloqueio = ""
        
        codDesbloqueio = GeraCodigoDesbloqueio(rsClientesMarcados!CNPJ, rsClientesMarcados!Razao, rsClientesMarcados!Mes_Referente, rsClientesMarcados!Ano_Referente)
        'Salva código Desbloqueio gerado
        If Not Vazio(codDesbloqueio) Then
           sSQL = "UPDATE empresas_desbloueio SET Cod_Desbloqueio = '" & codDesbloqueio & "' WHERE CNPJ = '" & rsClientesMarcados!CNPJ & "'"
           dbData.Execute sSQL
        End If
        
        rsClientesMarcados.MoveNext
    Loop
    
    Set rsClientesMarcados = Nothing
    
    Exit Sub
    
    Resume
    
deuErro:
    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
    Err.Clear
    Set rsClientesMarcados = Nothing
End Sub

Private Sub cmdGerarMes_Click()
   dbData.Execute "UPDATE empresas_desbloueio SET Mes_Referente = '" & cboMes.Text & "', Ano_Referente = '" & cboAno.Text & "' where MARCADO = 1;"
End Sub

Private Sub cmdGerenciaNet_Click()
If cboMes.Text = "" Then Exit Sub
If cboAno.Text = "" Then Exit Sub
Dim arqCSV As String, strArquivo As String, codDesbloqueio As String
Dim i As Long
Dim pJSON As Object
   
   On Error GoTo deuErro
   
   CommonDialog1.Filter = "Arquivo json(*.json)|*.json|Arquivo CSV(*.csv)|*.csv"
   CommonDialog1.ShowOpen
   arqCSV = CommonDialog1.FileName
   
   If (CommonDialog1.FileName = "") Then Exit Sub

   Set p = JSON.parse2(ReadTextFile(arqCSV))
   If Not (p Is Nothing) Then
      If JSON.GetParserErrors <> "" Then
         MsgBox JSON.GetParserErrors, vbCritical + vbOKOnly, "Parsing Error(s) occured"
      Else
         MsgBox "Foi encontrado " & p.Count & " título(s) no arquivo!", vbInformation + vbOKOnly
         For i = 1 To p.Count
            strArquivo = "CNPJ: " & p.Item(i).Item("customer").Item("document") & vbNewLine
            strArquivo = strArquivo & "Vencimento: " & p.Item(i).Item("expire_at") & vbNewLine
            strArquivo = strArquivo & "Valor: " & p.Item(i).Item("total") & vbNewLine
            strArquivo = strArquivo & "Valor Pago: " & p.Item(i).Item("paid_value")
            'MsgBox strArquivo, vbInformation
            'Caso localize o título vai marcar ele e salvar mes/ano referencia
            sSQL = "UPDATE empresas_desbloueio SET Mes_Referente = '" & cboMes.Text & "', Ano_Referente = '" & cboAno.Text & "', marcado = 1 WHERE CNPJ = '" & Format(p.Item(i).Item("customer").Item("document"), "@@.@@@.@@@/@@@@-@@") & "'"
            dbData.Execute sSQL
            'Gera código Desbloqueio
            sSQL = "SELECT RAZAO, CNPJ, celular, Mes_Referente, Ano_Referente, Cod_Desbloqueio FROM empresas_desbloueio WHERE CNPJ = '" & Format(p.Item(i).Item("customer").Item("document"), "@@.@@@.@@@/@@@@-@@") & "'"
            codDesbloqueio = GeraCodigoDesbloqueio(SQLExecutaRetorno(sSQL, "CNPJ", ""), SQLExecutaRetorno(sSQL, "RAZAO", ""), SQLExecutaRetorno(sSQL, "Mes_Referente", ""), SQLExecutaRetorno(sSQL, "Ano_Referente", ""))
            'Salva código Desbloqueio gerado
            sSQL = "UPDATE empresas_desbloueio SET Cod_Desbloqueio = '" & codDesbloqueio & "' WHERE CNPJ = '" & Format(p.Item(i).Item("customer").Item("document"), "@@.@@@.@@@/@@@@-@@") & "'"
            dbData.Execute sSQL
            'Envia mensagem com o código de Desbloqueio para o celular cadastrado
            iRetorno = WhatsAppEnviar(SQLExecutaRetorno(sSQL, "CNPJ", ""), SQLExecutaRetorno(sSQL, "RAZAO", ""), SQLExecutaRetorno(sSQL, "celular", ""), SQLExecutaRetorno(sSQL, "Cod_Desbloqueio", ""))
            DoEvents
         Next
         MostrarEmpresa
      End If
   Else
      MsgBox "Erro ao ler o arquivo " & CommonDialog1.FileName, vbCritical + vbOKOnly
   End If
   
   Exit Sub
   
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: LerAquivoGerenciaNet"
   Err.Clear
End Sub

Private Function GeraCodigoDesbloqueio(CNPJ As String, RazaoSocial As String, MesRef As String, AnoRef As String) As String
Dim codigoGerado As String, vCodDesbloqueio As String, vCodDesbTemp As String
Dim vCnpj As Integer, vQuantRazao As Integer, vNumeroMes As Integer, vDataBloqueio As String
Dim vDataInicio As Date, vDia As Integer, vMes As Integer, vMesInt As String, vAno As Integer, vMesRef As String

   On Error GoTo deuErro

    vCnpj = SomarDigitos(CNPJ)
    vQuantRazao = Len(RazaoSocial)
    
    If Vazio(vCnpj) Then Exit Function

    Select Case MesRef
        Case "Janeiro"
            vNumeroMes = 1
        Case "Fevereiro"
            vNumeroMes = 2
        Case "Março"
            vNumeroMes = 3
        Case "Abril"
            vNumeroMes = 4
        Case "Maio"
            vNumeroMes = 5
        Case "Junho"
            vNumeroMes = 6
        Case "Julho"
            vNumeroMes = 7
        Case "Agosto"
            vNumeroMes = 8
        Case "Setembro"
            vNumeroMes = 9
        Case "Outubro"
            vNumeroMes = 10
        Case "Novembro"
            vNumeroMes = 11
        Case "Dezembro"
            vNumeroMes = 12
    End Select

    'vDia = 30
    vMes = vNumeroMes
    vAno = AnoRef

    If vMes = 2 Then
        If AnoBisexto(Year(Date)) Then
            vDia = 29
        Else
            vDia = 28
        End If
    Else
        vDia = 30
    End If

    'Autonumeracao_Pagamentos
    vDataInicio = vDia & " / " & vMes & " / " & vAno
    vMesInt = Format(vDataInicio, "mmmm")
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
    
    GeraCodigoDesbloqueio$ = vCodDesbloqueio

   Exit Function
   
   Resume
   
deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: GeraCodigoDesbloqueio"
   Err.Clear
End Function

Private Function WhatsAppEnviar(CNPJ As String, RazaoSocial As String, Telefone As String, CodigoDesbloqueio As String) As Boolean
Dim iRetorno As Boolean, mensagemRetorno As String, mensagemErro As String, mensagemEnvio As String
Dim idmsg As Integer, codTelefone As Long, nTelefone As String, ComandoSQL As String, vUltLetra As String, vMensagemCodigo As String
Dim ekZap As zapzap.cZap

    On Error GoTo deuErro
    
    Set ekZap = New zapzap.cZap
    
    Call ekZap.CarregarConfiguracoes
    
    iRetorno = ekZap.setarDelay(5, mensagemRetorno, mensagemErro)
    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
       GoTo Caifora
    Else
    
    End If

    DoEvents
    
    lblStatus.Caption = ""
    DoEvents
    
    nTelefone = Telefone
    
    If Vazio(nTelefone) Then
       lblStatus.Caption = "O Cliente não tem nenhum telefone celular autorizado para envio de mensagem!"
       DoEvents
       GoTo PulaProximo
    End If
    
    nTelefone = "+55" + Retira(nTelefone, "()- ", UM_A_UM)
        
    idmsg = Int((Rnd * 999) + 1)
    lblStatus.Caption = "Enviando mensagem para " & Trim(RazaoSocial) & ". Aguarde..."
    DoEvents
    If Len(CodigoDesbloqueio) > 0 Then
       vUltLetra = Right(CodigoDesbloqueio, 1)
       vMensagemCodigo = "[MENSAGEM AUTOMÁTICA]: Seu código de desbloqueio é: " & CodigoDesbloqueio & "   - Obs: O último caractere é a letra '" & vUltLetra & "'. "
    End If
    mensagemEnvio = txtMensagem.Text
    mensagemEnvio = Substitui(mensagemEnvio, "{cliente}", Trim(RazaoSocial), SO_UM)
    mensagemEnvio = Substitui(mensagemEnvio, "{codigodesbloqueio}", Trim(vMensagemCodigo), SO_UM)
    ekZap.FTPurl = "ftp.ekklesiasoft.com.br"
    ekZap.FTPuser = "onlineinfo@ekklesiasoft.com.br"
    ekZap.FTPpassw = "Webier@online"
    ekZap.HTTPurl = "http://www.ekklesiasoft.com.br/zapzap/onlineinfo"
    iRetorno = ekZap.ConfigurarDLL("onlineinfo@ekklesiasoft.com.br", "6d3a55c40463afc5f5824e031bb0d0b445907", "6464", eProvedor_Solutek, mensagemErro)
   'iRetorno = ekZap.ConfigurarDLL("onlineinfo@ekklesiasoft.com.br", "65cce637c50487a7e4bee5203bb9320b709088", "6373", mensagemErro)

    DoEvents
    iRetorno = ekZap.Enviar(idmsg, nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, mensagemRetorno, mensagemErro)

    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
    Else
       If mensagemRetorno <> "" Then
          lblStatus.Caption = "Mensagem enviada para " & Trim(RazaoSocial)
          DoEvents
          'MsgBox mensagemRetorno, vbInformation + vbOKOnly
          ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
          mensagemErro$ = SQLExecuta(ComandoSQL)
          If Not Vazio(mensagemErro$) Then
             MsgBox mensagemErro$, vbCritical + vbOKOnly, "ERRO: Registrar Envio"
             mensagemErro = ""
          End If
       Else
          MsgBox mensagemErro, vbCritical + vbOKOnly
       End If
    End If
    Sleep 5000
PulaProximo:
    
    iRetorno = ekZap.setarDelay(3, mensagemRetorno, mensagemErro)
    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
       GoTo Caifora
    End If
    
    'lblStatus.Caption = "Envio de mensagem concluído com sucesso!"
    'DoEvents
    
    'ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
    'SQLExecuta ComandoSQL
    'DoEvents
    
    Set ekZap = Nothing
    
    WhatsAppEnviar = True
    
    DoEvents
    
    Exit Function
    
Caifora:
    MsgBox "O Cliente não tem nenhum telefone celular autorizado para envio de mensagem!" & vbCrLf & "Verifique no cadastro do cliente e tente novamente.", vbExclamation + vbOKOnly
    Set ekZap = Nothing
    DoEvents
    Exit Function
    
deuErro:
    lblStatus.Caption = Err.Description
    DoEvents
    Set ekZap = Nothing
    MsgBox Err.Description, vbCritical + vbOKOnly
    Err.Clear
    DoEvents
End Function

Private Function WhatsAppEnviarCodigo(CNPJ As String, RazaoSocial As String, Telefone As String, CodigoDesbloqueio As String) As Boolean
Dim iRetorno As Boolean, mensagemRetorno As String, mensagemErro As String, mensagemEnvio As String
Dim idmsg As Integer, codTelefone As Long, nTelefone As String, ComandoSQL As String, vUltLetra As String, vMensagemCodigo As String
Dim ekZap As zapzap.cZap

    On Error GoTo deuErro
    
    Set ekZap = New zapzap.cZap
    
    Call ekZap.CarregarConfiguracoes
    
    iRetorno = ekZap.setarDelay(5, mensagemRetorno, mensagemErro)
    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
       GoTo Caifora
    Else
    
    End If

    DoEvents
    
    lblStatus.Caption = ""
    DoEvents
    
    nTelefone = Telefone
    
    If Vazio(nTelefone) Then
       lblStatus.Caption = "O Cliente não tem nenhum telefone celular autorizado para envio de mensagem!"
       DoEvents
       GoTo PulaProximo
    End If
    
    nTelefone = "+55" + Retira(nTelefone, "()- ", UM_A_UM)
        
    idmsg = Int((Rnd * 999) + 1)
    lblStatus.Caption = "Enviando mensagem para " & Trim(RazaoSocial) & ". Aguarde..."
    DoEvents
    If Len(CodigoDesbloqueio) > 0 Then
       vUltLetra = Right(CodigoDesbloqueio, 1)
       vMensagemCodigo = "Seu código de desbloqueio é: " & CodigoDesbloqueio & " referente ao mês: " & vMesRef & "/" & vAnoRef & vbNewLine & "Use-o quando seu sistema bloquear!" & vbNewLine & "Obs: O último caractere do código é a letra " & vUltLetra & ". "
       'vMensagemCodigo = "Seu código de desbloqueio é: " & CodigoDesbloqueio & " referente ao mês: " & Mes_Referente & "Junho/2023" & Ano_Referente & vbNewLine & "Use-o quando seu sistema bloquear!" & vbNewLine & "Obs: O último caractere é a letra " & vUltLetra & ". "
    End If
    mensagemEnvio = "Olá " & Trim(RazaoSocial) & ", Tudo bem?"
    mensagemEnvio = mensagemEnvio & vbNewLine & Trim(vMensagemCodigo)
    ekZap.FTPurl = "ftp.ekklesiasoft.com.br"
    ekZap.FTPuser = "onlineinfo@ekklesiasoft.com.br"
    ekZap.FTPpassw = "Webier@online"
    ekZap.HTTPurl = "http://www.ekklesiasoft.com.br/zapzap/onlineinfo"
    iRetorno = ekZap.ConfigurarDLL("onlineinfo@ekklesiasoft.com.br", "35d757ce6f23323ff688299b67029f62435839", "6713", eProvedor_Solutek, mensagemErro)
    DoEvents
    iRetorno = ekZap.Enviar(idmsg, nTelefone, mensagemEnvio, mensagemRetorno, mensagemErro)
'    iRetorno = ekZap.ConfigurarDLL("webieronline@gmail.com", "df925a69bcfbca2ae8465cf40e869fa8519253", "5878", mensagemErro)

    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
    Else
       If mensagemRetorno <> "" Then
          lblStatus.Caption = "Mensagem enviada para " & Trim(RazaoSocial)
          DoEvents
          'MsgBox mensagemRetorno, vbInformation + vbOKOnly
          ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
          mensagemErro$ = SQLExecuta(ComandoSQL)
          If Not Vazio(mensagemErro$) Then
             MsgBox mensagemErro$, vbCritical + vbOKOnly, "ERRO: Registrar Envio"
             mensagemErro = ""
          End If
       Else
          MsgBox mensagemErro, vbCritical + vbOKOnly
       End If
    End If
    Sleep 5000
PulaProximo:
    
    iRetorno = ekZap.setarDelay(3, mensagemRetorno, mensagemErro)
    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
       GoTo Caifora
    End If
    
    'lblStatus.Caption = "Envio de mensagem concluído com sucesso!"
    'DoEvents
    
    'ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
    'SQLExecuta ComandoSQL
    'DoEvents
    
    Set ekZap = Nothing
    
    WhatsAppEnviarCodigo = True
    
    DoEvents
    
    Exit Function
    
Caifora:
    MsgBox "O Cliente não tem nenhum telefone celular autorizado para envio de mensagem!" & vbCrLf & "Verifique no cadastro do cliente e tente novamente.", vbExclamation + vbOKOnly
    Set ekZap = Nothing
    DoEvents
    Exit Function
    
deuErro:
    lblStatus.Caption = Err.Description
    DoEvents
    Set ekZap = Nothing
    MsgBox Err.Description, vbCritical + vbOKOnly
    Err.Clear
    DoEvents
End Function

Private Function WhatsAppEnviarAnexo(CNPJ As String, RazaoSocial As String, Telefone As String) As Boolean
Dim iRetorno As Boolean, mensagemRetorno As String, mensagemErro As String, mensagemEnvio As String
Dim idmsg As Integer, codTelefone As Long, nTelefone As String, ComandoSQL As String
Dim ekZap As zapzap.cZap

    On Error GoTo deuErro
    
    Set ekZap = New zapzap.cZap
    
    Call ekZap.CarregarConfiguracoes
    
    iRetorno = ekZap.setarDelay(5, mensagemRetorno, mensagemErro)
    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
       GoTo Caifora
    Else
    
    End If

    DoEvents
    
    lblStatus.Caption = ""
    DoEvents
    
    nTelefone = Telefone
    
    If Vazio(nTelefone) Then
       lblStatus.Caption = "O Cliente não tem nenhum telefone celular autorizado para envio de mensagem!"
       DoEvents
       GoTo PulaProximo
    End If
    
    nTelefone = "+55" + Retira(nTelefone, "()- ", UM_A_UM)
        
    idmsg = Int((Rnd * 999) + 1)
    lblStatus.Caption = "Enviando mensagem para " & Trim(RazaoSocial) & ". Aguarde..."
    DoEvents
    mensagemEnvio = txtMensagem.Text
    mensagemEnvio = Substitui(mensagemEnvio, "{cliente}", Trim(RazaoSocial), SO_UM)
    ekZap.FTPurl = "ftp.ekklesiasoft.com.br"
    ekZap.FTPuser = "onlineinfo@ekklesiasoft.com.br"
    ekZap.FTPpassw = "Webier@online"
    ekZap.HTTPurl = "http://www.ekklesiasoft.com.br/zapzap/onlineinfo/"
    iRetorno = ekZap.ConfigurarDLL("webieronline@gmail.com", "cd2582694794062678efd0f9157bd299986371", "6086", eProvedor_Solutek, mensagemErro)
    DoEvents
    If Right(txtCaminho.Text, 3) = "jpg" Or Right(txtCaminho.Text, 4) = "jpeg" Then
       iRetorno = ekZap.enviarimagem(idmsg, nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, txtCaminho.Text, mensagemRetorno, mensagemErro)
    Else
       iRetorno = ekZap.enviararquivo(idmsg, nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, txtCaminho.Text, mensagemRetorno, mensagemErro)
    End If

    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
    Else
       If mensagemRetorno <> "" Then
          lblStatus.Caption = "Mensagem enviada para " & Trim(RazaoSocial)
          DoEvents
          'MsgBox mensagemRetorno, vbInformation + vbOKOnly
          ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
          mensagemErro$ = SQLExecuta(ComandoSQL)
          If Not Vazio(mensagemErro$) Then
             MsgBox mensagemErro$, vbCritical + vbOKOnly, "ERRO: Registrar Envio"
             mensagemErro = ""
          End If
       Else
          MsgBox mensagemErro, vbCritical + vbOKOnly
       End If
    End If
    Sleep 5000
PulaProximo:
    
    iRetorno = ekZap.setarDelay(3, mensagemRetorno, mensagemErro)
    If Not iRetorno Then
       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
       GoTo Caifora
    End If
    
    'lblStatus.Caption = "Envio de mensagem concluído com sucesso!"
    'DoEvents
    
    'ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
    'SQLExecuta ComandoSQL
    'DoEvents
    
    Set ekZap = Nothing
    
    WhatsAppEnviarAnexo = True
    
    DoEvents
    
    Exit Function
    
Caifora:
    MsgBox "O Cliente não tem nenhum telefone celular autorizado para envio de mensagem!" & vbCrLf & "Verifique no cadastro do cliente e tente novamente.", vbExclamation + vbOKOnly
    Set ekZap = Nothing
    DoEvents
    Exit Function
    
deuErro:
    lblStatus.Caption = Err.Description
    DoEvents
    Set ekZap = Nothing
    MsgBox Err.Description, vbCritical + vbOKOnly
    Err.Clear
    DoEvents
End Function

Private Sub cmdLocalizarArquivo_Click()
Dim FSys As FileSystemObject 'referencia que nao deixa copiar arquivos duplicados (PROJECT / REFERENCES e selecionar MICROSOFT SCRIPTING RUNTIME)
    Set FSys = New FileSystemObject
    
    'CommonDialog1.Filter = "Imagens JPG(*.jpg)|*.jpg"
    CommonDialog1.ShowOpen
    txtCaminho.Text = CommonDialog1.FileName
    
    If (CommonDialog1.FileName = "") Then Exit Sub
    
    'If Not FSys.FileExists(Caminho & CommonDialog1.FileTitle) Then  'se o arquivo nao existir na pasta ele copia
    '   FileCopy txtCaminho.Text, Caminho & CommonDialog1.FileTitle
    'End If
       'txtCaminho.Text = Caminho & CommonDialog1.FileTitle
    txtCaminho.Text = CommonDialog1.FileName
       'picLogo.Picture = LoadPicture(txtCaminho.Text) 'mostrar a imagem
    Set FSys = Nothing
End Sub

Private Sub cmdMostrarTudo_Click()
    MostrarEmpresa
End Sub

Private Sub Form_Load()
    Set moCombo = New cComboHelper
    MostrarEmpresa
    
    If WhatsAppConectado Then
       imgOnLine.Visible = True
       imgOFFLine.Visible = False
       DoEvents
    Else
       imgOnLine.Visible = False
       imgOFFLine.Visible = True
       DoEvents
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
    LimparEmpresa
    txtFantasia.Text = (Grid.TextMatrix(Grid.Row, 1))
    txtRazao.Text = (Grid.TextMatrix(Grid.Row, 2))
    mskCPF.Text = (Grid.TextMatrix(Grid.Row, 3))
    mskCelular.Text = (Grid.TextMatrix(Grid.Row, 6))
    cboCidade.Text = (Grid.TextMatrix(Grid.Row, 8))
    cboEstado.Text = (Grid.TextMatrix(Grid.Row, 9))
    If (Grid.TextMatrix(Grid.Row, 10)) = "NÃO" Then
        optCNPJ.Value = True
        optCPF.Value = False
    Else
        optCNPJ.Value = False
        optCPF.Value = True
    End If
    
    txtCodDesbloqueio.Text = ""
    txtCodDesbloqueioTemp.Text = ""
End Sub

Private Sub mskCelular_KeyPress(KeyAscii As Integer)
    mskCelular.Mask = "(##) #####-####"
End Sub

Private Sub mskCPF_KeyPress(KeyAscii As Integer)
    If optCNPJ.Value = False Then
        mskCPF.Mask = "###.###.###-##"
    Else
        mskCPF.Mask = "##.###.###/####-##"
    End If
End Sub

Private Sub chameleonButton1_Click()
    Clipboard.Clear
    Clipboard.SetText mskCPF.Text
End Sub

Private Sub cmdAdicionar_Click()
    If txtFantasia.Text = "" Or txtRazao.Text = "" Or mskCPF.Text = "" Then Exit Sub
    
    sSQL = "SELECT CNPJ FROM  empresas_desbloueio WHERE CNPJ = '" & mskCPF.Text & "';"
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

Private Function Inserir_Dados() As Boolean
Dim vNovoCodigo As Integer

    'autonumeração
    sSQL = "SELECT MAX(CODIGO) r FROM empresas_desbloueio "
    vNovoCodigo = SQLExecutaRetorno(sSQL, "r", 0) + 1
    
    'Comando de inclusão
    sSQL = "INSERT INTO empresas_desbloueio (" & _
       "fantasia, razao, cnpj, celular, codigo, marcado, cidade, estado, CPF) VALUES ('" & _
       txtFantasia.Text & "', '" & txtRazao.Text & "', '" & mskCPF.Text & "', '" & mskCelular.Text & "', " & vNovoCodigo & ", 0, '" & cboCidade.Text & "', '" & cboEstado.Text & "', '" & Abs(optCPF.Value) & "')"
    
    'Retorna o resultado da atualização
    Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdNovo_Click()
    LimparEmpresa
    MostrarEmpresa
End Sub

Private Sub LimparEmpresa()
    txtFantasia.Text = ""
    txtRazao.Text = ""
    cboCidade.Text = ""
    cboEstado.Text = ""
    mskCPF.Mask = ""
    mskCPF.Text = ""
    mskCelular.Mask = ""
    mskCelular.Text = ""
End Sub

Private Sub MostrarEmpresa()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado, (CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF FROM  empresas_desbloueio  ORDER BY FANTASIA;"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    sSQL = "SELECT * FROM  empresas_desbloueio where marcado = 1;"
    Set r = dbData.OpenRecordset(sSQL)
    
    lblMarcado.Caption = r.RecordCount
End Sub
Private Sub cmdLocalizar_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado,(CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF  FROM  empresas_desbloueio where (FANTASIA LIKE '%" & txtFantasia.Text & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim x As Integer

    With Grid
       .Clear
       .Cols = 11
       .rows = 2
       
       .ColWidth(0) = 0
       .ColWidth(1) = 2500
       .ColWidth(2) = 4000
       .ColWidth(3) = 1800
       .ColWidth(4) = 0
       .ColWidth(5) = 0
       .ColWidth(6) = 1400
       .ColWidth(7) = 0
       .ColWidth(8) = 1400
       .ColWidth(9) = 600
       .ColWidth(10) = 600
       
       
       For x = 0 To .Cols - 1
          .Col = x
          .Row = 0
          .CellFontBold = True
       Next
       
       .TextMatrix(0, 1) = "FANTASIA"
       .TextMatrix(0, 2) = "RAZÃO."
       .TextMatrix(0, 3) = "CNPJ"
       .TextMatrix(0, 4) = "CODIGO"
       .TextMatrix(0, 6) = "CELULAR"
       .TextMatrix(0, 8) = "CIDADE"
       .TextMatrix(0, 9) = "UF"
       .TextMatrix(0, 10) = "CPF"
       
       .Redraw = False
       
       i = 1
       If Not rTabela Is Nothing Then
          Do While Not rTabela.EOF
             .TextMatrix(.rows - 1, 1) = rTabela("FANTASIA")
             .TextMatrix(.rows - 1, 2) = rTabela("RAZAO")
             .TextMatrix(.rows - 1, 3) = rTabela("CNPJ")
             .TextMatrix(.rows - 1, 4) = rTabela("CODIGO")
             .TextMatrix(.rows - 1, 5) = rTabela("vMarcado")
             .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("CELULAR"))
             .TextMatrix(.rows - 1, 7) = rTabela("vEnviado")
             .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("CIDADE"))
             .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("ESTADO"))
             .TextMatrix(.rows - 1, 10) = rTabela("vCPF")
             rTabela.MoveNext
             
             .rows = .rows + 1
             i = i + 1
          Loop
       End If
       
       
       For i = 1 To .rows - 1
           For j = 0 To .Cols - 1
              .Col = j
              .Row = i
        
              If .TextMatrix(i, 5) = "NÃO" And .TextMatrix(i, 7) = "NÃO" Then
                 .CellForeColor = vbBlack
              ElseIf .TextMatrix(i, 5) = "SIM" And .TextMatrix(i, 7) = "NÃO" Then
                 .CellForeColor = vbRed
              ElseIf .TextMatrix(i, 5) = "SIM" And .TextMatrix(i, 7) = "SIM" Then
                 .CellForeColor = vbBlue
              Else
                 .CellForeColor = vbBlack
              End If
              
           Next
        Next
       
       .rows = .rows - 1
       .Redraw = True
    End With
End Sub

Private Sub cmdMarcar_Click()
    i = Grid.Row
    dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 1 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
    MostrarEmpresa
    'sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado FROM  empresas_desbloueio ORDER BY RAZAO;"
    'Set r = dbData.OpenRecordset(sSQL)
    'FormatarGrid r
End Sub

Private Sub cmdDesmarcar_Click()
    i = Grid.Row
    dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 0 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
    MostrarEmpresa
End Sub

Private Sub cmdDesmarcarTodos_Click()
    dbData.Execute "UPDATE empresas_desbloueio SET MARCADO = 0,  Mes_Referente = '', Ano_Referente = '', Cod_Desbloqueio = '', Enviado = 0;"
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
Dim vCnpj As Integer
Dim vQuantRazao As Integer
Dim vNumeroMes As Integer
Dim vDataInicio As Date
Dim vDia As Integer
Dim vMes As Integer
Dim vMesInt As String
Dim vAno As Integer
Dim vMesRef As String
    
    If txtFantasia.Text = "" Or txtRazao.Text = "" Or mskCPF.Text = "" Then Exit Sub
    If cboMes.Text = "" Then Exit Sub
    If cboAno.Text = "" Then Exit Sub

    vCnpj = SomarDigitos(mskCPF.Text)
    vQuantRazao = Len(txtRazao.Text)
    
    If cboMes.Text = "Janeiro" Then
        vNumeroMes = 1
    ElseIf cboMes.Text = "Fevereiro" Then
        vNumeroMes = 2
    ElseIf cboMes.Text = "Março" Then
        vNumeroMes = 3
    ElseIf cboMes.Text = "Abril" Then
        vNumeroMes = 4
    ElseIf cboMes.Text = "Maio" Then
        vNumeroMes = 5
    ElseIf cboMes.Text = "Junho" Then
        vNumeroMes = 6
    ElseIf cboMes.Text = "Julho" Then
        vNumeroMes = 7
    ElseIf cboMes.Text = "Agosto" Then
        vNumeroMes = 8
    ElseIf cboMes.Text = "Setembro" Then
        vNumeroMes = 9
    ElseIf cboMes.Text = "Outubro" Then
        vNumeroMes = 10
    ElseIf cboMes.Text = "Novembro" Then
        vNumeroMes = 11
    ElseIf cboMes.Text = "Dezembro" Then
        vNumeroMes = 12
    End If

    'começa a criação
    'vDia = 30
    vMes = vNumeroMes
    vAno = cboAno
    
    If vMes = 2 Then
        If AnoBisexto(Year(Date)) Then
            vDia = 29
        Else
            vDia = 28
        End If
    Else
        vDia = 30
    End If
    
    Dim vDataBloqueio As String
    
    'Autonumeracao_Pagamentos
    
    'If chkProximo.Value = 1 Then
    '    vDataInicio = vDia & " / " & vMes & " / " & vAno
    '    vDataInicio = Format(DateAdd("m", Val(1), vDataInicio), "dd/mm/yy")
    '    vMesInt = Format(vDataInicio, "mmmm")
    '    vAno = Year(vDataInicio)
    '    vMesRef = vMesInt & "/" & vAno
    'Else
        vDataInicio = vDia & " / " & vMes & " / " & vAno
        vMesInt = Format(vDataInicio, "mmmm")
    '    vAno = Year(vDataInicio)
    '    vMesRef = vMesInt & "/" & vAno
    'End If
    
    'vDataBloqueio = Format(DateAdd("d", Val(5), vDataInicio), "dd/mm/yy")
    
    'codigo de desbloqueio
        
    '    If vMesInt = "janeiro" Then
    '        vNumeroMes = 1
    '    ElseIf vMesInt = "fevereiro" Then
    '        vNumeroMes = 2
    '    ElseIf vMesInt = "março" Then
    '        vNumeroMes = 3
    '    ElseIf vMesInt = "abril" Then
    '        vNumeroMes = 4
    '    ElseIf vMesInt = "maio" Then
    '        vNumeroMes = 5
    '    ElseIf vMesInt = "junho" Then
    '        vNumeroMes = 6
    '    ElseIf vMesInt = "julho" Then
    '        vNumeroMes = 7
    '    ElseIf vMesInt = "agosto" Then
    '        vNumeroMes = 8
    '    ElseIf vMesInt = "setembro" Then
    '        vNumeroMes = 9
    '    ElseIf vMesInt = "outubro" Then
    '        vNumeroMes = 10
    '    ElseIf vMesInt = "novembro" Then
     '       vNumeroMes = 11
     '   ElseIf vMesInt = "dezembro" Then
    '        vNumeroMes = 12
    '    End If
        
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
        
    txtCodDesbloqueio.Text = vCodDesbloqueio
    txtCodDesbloqueioTemp.Text = vCodDesbTemp
End Sub

Private Function AnoBisexto(ValAno As Single) As Boolean
    If (ValAno Mod 4 = 0) And ((ValAno Mod 100 <> 0) Or (ValAno Mod 400 = 0)) Then
            AnoBisexto = True
        Else
            AnoBisexto = False
    End If
End Function

Public Function SomarDigitos(CNPJ As String) As Integer
    Dim s As Integer
    Dim i As Integer
    For i = 1 To Len(CNPJ)
      If IsNumeric(Mid(CNPJ, i, 1)) Then
        s = s + Mid(CNPJ, i, 1)
      End If
    Next
    SomarDigitos = s
End Function

Private Sub cmdPrepara_Click()
Dim vUltLetra As String
Dim vMsgID As String
Dim vMsgCode As String
Dim vMsgOBS As String

    If txtCodDesbloqueio.Text = "" Then Exit Sub
    
    vUltLetra = Right(txtCodDesbloqueio.Text, 1)
    vMsgID = "Olá " & Trim(txtFantasia) & ", Tudo bem?"
    vMsgCode = "Seu código de desbloqueio é: " & txtCodDesbloqueio.Text & " referente ao mês: " & cboMes.Text & "/" & cboAno.Text & ". Use-o quando seu sistema bloquear!"
    vMsgOBS = "Obs: O último caractere do código é a letra " & vUltLetra & ". "
    
    Clipboard.Clear
    Clipboard.SetText vMsgID & vbNewLine & Trim(vMsgCode) & vbNewLine & vMsgOBS
    'Clipboard.SetText "[MENSAGEM AUTOMÁTICA]: Seu código de desbloqueio é: " & txtCodDesbloqueio.Text & " referente ao mês: Julho/2023 " & vbNewLine & " Obs: O último caractere é a letra '" & vUltLetra & "'. "
End Sub

Private Sub cmdPrepara2_Click()
    Clipboard.Clear
    Clipboard.SetText "Seu código de desbloqueio temporário é: " & txtCodDesbloqueioTemp.Text
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazao_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
