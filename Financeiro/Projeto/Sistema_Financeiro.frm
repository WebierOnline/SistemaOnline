VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Sistema_Financeiro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADMINISTRATIVO"
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
      Tab             =   1
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
      TabCaption(0)   =   "DESBLOQUEIO"
      TabPicture(0)   =   "Sistema_Financeiro.frx":23D2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblMarcado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStatus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAnual"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdMostrarMarcados"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdImprimir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGerenciaNet"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdMarcarTodos"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdEnviarTodos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ccmdMostrarRazao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdMostrarTudo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdEnviarUm"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdDesmarcarTodos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdDesmarcar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdMarcar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Grid"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "frmCadastro"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "IBPT"
      TabPicture(1)   =   "Sistema_Financeiro.frx":23EE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ALÍQUOTAS CBS/IBS"
      TabPicture(2)   =   "Sistema_Financeiro.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraAlq"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraAlq 
         Caption         =   "Alíquotas da Reforma Tributária (CBS / IBS) - publicadas para os clientes pela nuvem"
         Height          =   6120
         Left            =   -74880
         TabIndex        =   75
         Top             =   480
         Width           =   11760
         Begin VB.Label lblAlqTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   330
            Width           =   345
         End
         Begin VB.ComboBox cboAlqTipo 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label lblAlqEscopo 
            AutoSize        =   -1  'True
            Caption         =   "Escopo:"
            Height          =   195
            Left            =   2520
            TabIndex        =   2
            Top             =   330
            Width           =   525
         End
         Begin VB.TextBox txtAlqEscopo 
            Height          =   288
            Left            =   3180
            MaxLength       =   7
            TabIndex        =   3
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label lblAlqEscopoDica 
            AutoSize        =   -1  'True
            Caption         =   "TODOS, UF (ex.: SP) ou código IBGE do município (7 dígitos)"
            Height          =   195
            Left            =   4560
            TabIndex        =   4
            Top             =   330
            Width           =   4200
         End
         Begin VB.Label lblAlqAliq 
            AutoSize        =   -1  'True
            Caption         =   "Alíq. %:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   780
            Width           =   585
         End
         Begin VB.TextBox txtAlqAliq 
            Height          =   288
            Left            =   840
            MaxLength       =   8
            TabIndex        =   6
            Top             =   735
            Width           =   900
         End
         Begin VB.Label lblAlqIni 
            AutoSize        =   -1  'True
            Caption         =   "Início:"
            Height          =   195
            Left            =   2040
            TabIndex        =   7
            Top             =   780
            Width           =   465
         End
         Begin VB.TextBox txtAlqIni 
            Height          =   288
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   8
            Top             =   735
            Width           =   1200
         End
         Begin VB.Label lblAlqFim 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
            Height          =   195
            Left            =   4140
            TabIndex        =   9
            Top             =   780
            Width           =   285
         End
         Begin VB.TextBox txtAlqFim 
            Height          =   288
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   10
            Top             =   735
            Width           =   1200
         End
         Begin VB.Label lblAlqObs 
            AutoSize        =   -1  'True
            Caption         =   "Obs.:"
            Height          =   195
            Left            =   6060
            TabIndex        =   11
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox txtAlqObs 
            Height          =   288
            Left            =   6540
            MaxLength       =   100
            TabIndex        =   12
            Top             =   735
            Width           =   4980
         End
         Begin VB.CommandButton cmdAlqIncluir 
            Caption         =   "Incluir / Alterar"
            Height          =   360
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   1800
         End
         Begin VB.CommandButton cmdAlqExcluir 
            Caption         =   "Excluir"
            Height          =   360
            Left            =   2040
            TabIndex        =   14
            Top             =   1200
            Width           =   1440
         End
         Begin VB.CommandButton cmdAlqLimpar 
            Caption         =   "Limpar"
            Height          =   360
            Left            =   3600
            TabIndex        =   15
            Top             =   1200
            Width           =   1440
         End
         Begin MSFlexGridLib.MSFlexGrid grdAlq 
            Height          =   3600
            Left            =   120
            TabIndex        =   16
            Top             =   1740
            Width           =   11520
            _ExtentX        =   20320
            _ExtentY        =   6350
            _Version        =   393216
         End
         Begin VB.CommandButton cmdAlqNuvem 
            Caption         =   "Carregar da nuvem"
            Height          =   360
            Left            =   120
            TabIndex        =   17
            Top             =   5520
            Width           =   2160
         End
         Begin VB.CommandButton cmdAlqEnviar 
            Caption         =   "Gerar e enviar"
            Height          =   360
            Left            =   2400
            TabIndex        =   18
            Top             =   5520
            Width           =   2160
         End
         Begin VB.Label lblAlqVersao 
            AutoSize        =   -1  'True
            Caption         =   ""
            Height          =   195
            Left            =   4800
            TabIndex        =   19
            Top             =   5580
            Width           =   6840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabela IBPT"
         Height          =   1035
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   7095
         Begin VB.TextBox txtLocalIBPT 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   780
            TabIndex        =   72
            Top             =   240
            Width           =   5715
         End
         Begin VB.CommandButton cmdBuscarIBPT 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   315
            Left            =   6540
            TabIndex        =   71
            Top             =   240
            Width           =   315
         End
         Begin ChamaleonBtn.chameleonButton cmdEnviarIBPT 
            Height          =   315
            Left            =   5760
            TabIndex        =   74
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Enviar"
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
            MICON           =   "Sistema_Financeiro.frx":2426
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblDirXML 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tabela:"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mensagens para clientes"
         Height          =   1875
         Left            =   -74940
         TabIndex        =   33
         Top             =   7620
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
            MICON           =   "Sistema_Financeiro.frx":2442
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
         Begin ChamaleonBtn.chameleonButton cmdConfigDLLWhats 
            Height          =   315
            Left            =   9300
            TabIndex        =   67
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Configurar DLL WhatsApp"
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
         Left            =   -74940
         TabIndex        =   15
         Top             =   6720
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
         Left            =   -74940
         TabIndex        =   12
         Top             =   4680
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
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   6482
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdMarcar 
         Height          =   315
         Left            =   -74880
         TabIndex        =   0
         Top             =   4080
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
         Left            =   -74040
         TabIndex        =   1
         Top             =   4080
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
         Left            =   -71820
         TabIndex        =   2
         Top             =   4080
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
         Left            =   -66480
         TabIndex        =   7
         Top             =   4080
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
         Left            =   -65760
         TabIndex        =   8
         Top             =   60
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
         Left            =   -64320
         TabIndex        =   9
         Top             =   60
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
         Left            =   -64860
         TabIndex        =   10
         Top             =   4080
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
         Left            =   -73140
         TabIndex        =   11
         Top             =   4080
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
         Left            =   -67560
         TabIndex        =   56
         Top             =   4080
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
         Left            =   -70320
         TabIndex        =   61
         Top             =   4080
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
      Begin ChamaleonBtn.chameleonButton cmdMostrarMarcados 
         Height          =   315
         Left            =   -67200
         TabIndex        =   68
         Top             =   60
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Mostrar Marcados"
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
         MICON           =   "Sistema_Financeiro.frx":45A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAnual 
         Height          =   315
         Left            =   -68040
         TabIndex        =   69
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Anual"
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
         MICON           =   "Sistema_Financeiro.frx":45C4
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
         Left            =   -74880
         TabIndex        =   57
         Top             =   4440
         Width           =   750
      End
      Begin VB.Label lblMarcado 
         AutoSize        =   -1  'True
         Caption         =   "0000"
         Height          =   195
         Left            =   -69480
         TabIndex        =   55
         Top             =   4140
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
Dim alqVersaoCarregada As String   'versao (AAAAMMDDHHMMSS) do arquivo de aliquotas que originou a grade

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Function EA(ByVal s As String) As String
    EA = Replace(s, "'", "''")
End Function

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
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado, (CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF FROM  empresas_desbloueio ORDER BY RAZAO;"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    sSQL = "SELECT * FROM  empresas_desbloueio where marcado = 1;"
    Set r = dbData.OpenRecordset(sSQL)
    
    lblMarcado.Caption = r.RecordCount
End Sub

Private Sub cmdMostrarMarcados_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado, (CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF FROM  empresas_desbloueio where marcado = 1 ORDER BY RAZAO;"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    sSQL = "SELECT * FROM  empresas_desbloueio where marcado = 1;"
    Set r = dbData.OpenRecordset(sSQL)
    
    lblMarcado.Caption = r.RecordCount
End Sub

Private Sub chameleonButton2_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado,(CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF  FROM  empresas_desbloueio where (RAZAO LIKE '%" & EA(txtRazao.TEXT) & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub chameleonButton3_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado, (CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF FROM  empresas_desbloueio where (CNPJ LIKE '%" & EA(mskCPF.TEXT) & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub chameleonButton4_Click()
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado,(CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF  FROM  empresas_desbloueio where (cidade LIKE '%" & EA(cboCidade.TEXT) & "%')"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub cmdConfigDLLWhats_Click()
   Call ZapConfigurar("onlineinfo@ekklesiasoft.com.br", "800d9cc25bda5c1b88efa6044774a021204192", "6755")
   DoEvents
   Call ZapStatus
End Sub

Private Sub cmdLerQRCode_Click()
Dim iRetorno As Boolean
    cmdLerQRCode.Enabled = False
    DoEvents
    Call WhatsAppConectar
    If WhatsAppConectado Then
       imgOnLine.Visible = True
       imgOFFLine.Visible = False
    Else
       imgOnLine.Visible = False
       imgOFFLine.Visible = True
    End If
    cmdLerQRCode.Enabled = True
    DoEvents
End Sub

Private Sub cmdDesconectarWhats_Click()
Dim iRetorno As Boolean
    cmdDesconectarWhats.Enabled = False
    DoEvents
    iRetorno = WhatsAppDesconectar
    If iRetorno Then
       imgOnLine.Visible = False
       imgOFFLine.Visible = True
    End If
    cmdDesconectarWhats.Enabled = True
    DoEvents
End Sub

Private Sub cmdEditar_Click()
    If txtFantasia.TEXT = "" Or txtRazao.TEXT = "" Or mskCPF.TEXT = "" Then Exit Sub
    
    If Not Atualizar_Dados Then
       ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
       Exit Sub
    End If
    
    'MostrarEmpresa
    LimparEmpresa
End Sub

Private Function Atualizar_Dados() As Boolean
i = Grid.Row

sSQL = "UPDATE empresas_desbloueio SET razao = '" & EA(txtRazao.TEXT) & "', fantasia = '" & EA(txtFantasia.TEXT) & "', cnpj = '" & EA(mskCPF.TEXT) & "', celular = '" & EA(mskCelular.TEXT) & "', cidade = '" & EA(cboCidade.TEXT) & "', estado = '" & EA(cboEstado.TEXT) & "', cpf = '" & Abs(optCPF.Value) & "' WHERE (codigo = " & Grid.TextMatrix(i, 4) & ");"

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
   dbData.Execute "UPDATE empresas_desbloueio SET Mes_Referente = '" & EA(cboMes.TEXT) & "', Ano_Referente = '" & EA(cboAno.TEXT) & "' where MARCADO = 1;"
End Sub

Private Function FormatarDocumento(ByVal doc As String) As String
    Dim soNumeros As String, k As Integer, c As String
    soNumeros = ""
    For k = 1 To Len(doc)
        c = Mid(doc, k, 1)
        If c >= "0" And c <= "9" Then soNumeros = soNumeros & c
    Next k
    Select Case Len(soNumeros)
        Case 11
            FormatarDocumento = Format(soNumeros, "@@@.@@@.@@@-@@")
        Case 14
            FormatarDocumento = Format(soNumeros, "@@.@@@.@@@/@@@@-@@")
        Case Else
            FormatarDocumento = doc
    End Select
End Function

Private Sub cmdGerenciaNet_Click()
If cboMes.TEXT = "" Then Exit Sub
If cboAno.TEXT = "" Then Exit Sub
Dim arqCSV As String, strArquivo As String
Dim vNaoEncontrados As String, vDocFormatado As String, vProcessados As String
Dim i As Long
Dim pJSON As Object
Dim rAnual As ADODB.Recordset

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
         vNaoEncontrados = ""
         vProcessados = ""
         For i = 1 To p.Count
            strArquivo = "CNPJ: " & p.item(i).item("customer").item("document") & vbNewLine
            strArquivo = strArquivo & "Vencimento: " & p.item(i).item("expire_at") & vbNewLine
            strArquivo = strArquivo & "Valor: " & p.item(i).item("total") & vbNewLine
            strArquivo = strArquivo & "Valor Pago: " & p.item(i).item("paid_value")
            vDocFormatado = FormatarDocumento(p.item(i).item("customer").item("document"))
            'MsgBox strArquivo, vbInformation
            If SQLExecutaRetorno("SELECT COUNT(*) AS Qtd FROM empresas_desbloueio WHERE CNPJ = '" & EA(vDocFormatado) & "'", "Qtd", 0) = 0 Then
               vNaoEncontrados = vNaoEncontrados & vDocFormatado & " - " & p.item(i).item("customer").item("name") & vbNewLine
            Else
               'Caso localize o título vai marcar, gerar codigo e enviar WhatsApp
               MarcarGerarEnviar vDocFormatado
               vProcessados = vProcessados & "|" & vDocFormatado & "|"
            End If
            DoEvents
         Next

         'Empresas com plano anual nao aparecem no relatorio mensal do GerenciaNet,
         'mas recebem o mesmo tratamento (marcar + gerar codigo + enviar WhatsApp)
         sSQL = "SELECT CNPJ FROM empresas_desbloueio WHERE anual = 1"
         Set rAnual = dbData.OpenRecordset(sSQL)
         Do While Not rAnual.EOF
            If InStr(1, vProcessados, "|" & rAnual("CNPJ") & "|") = 0 Then
               MarcarGerarEnviar rAnual("CNPJ")
            End If
            DoEvents
            rAnual.MoveNext
         Loop
         If rAnual.State <> 0 Then rAnual.Close
         Set rAnual = Nothing

         MostrarEmpresa
         If vNaoEncontrados <> "" Then
            MsgBox "As seguintes empresas do arquivo NAO foram encontradas no cadastro:" & vbNewLine & vbNewLine & vNaoEncontrados, vbExclamation, "Empresas nao cadastradas"
         End If
      End If
   Else
      MsgBox "Erro ao ler o arquivo " & CommonDialog1.FileName, vbCritical + vbOKOnly
   End If

   Exit Sub

deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: LerAquivoGerenciaNet"
   Err.Clear
End Sub

Private Sub MarcarGerarEnviar(ByVal vCnpj As String)
Dim codDesbloqueio As String
    'Marca a empresa e salva mes/ano referencia
    sSQL = "UPDATE empresas_desbloueio SET Mes_Referente = '" & EA(cboMes.TEXT) & "', Ano_Referente = '" & EA(cboAno.TEXT) & "', marcado = 1 WHERE CNPJ = '" & EA(vCnpj) & "'"
    dbData.Execute sSQL
    'Gera código Desbloqueio
    sSQL = "SELECT RAZAO, CNPJ, celular, Mes_Referente, Ano_Referente, Cod_Desbloqueio FROM empresas_desbloueio WHERE CNPJ = '" & EA(vCnpj) & "'"
    codDesbloqueio = GeraCodigoDesbloqueio(SQLExecutaRetorno(sSQL, "CNPJ", ""), SQLExecutaRetorno(sSQL, "RAZAO", ""), SQLExecutaRetorno(sSQL, "Mes_Referente", ""), SQLExecutaRetorno(sSQL, "Ano_Referente", ""))
    'Salva código Desbloqueio gerado
    sSQL = "UPDATE empresas_desbloueio SET Cod_Desbloqueio = '" & EA(codDesbloqueio) & "' WHERE CNPJ = '" & EA(vCnpj) & "'"
    dbData.Execute sSQL
    'Envia mensagem com o código de Desbloqueio para o celular cadastrado
    'sSQL = "SELECT RAZAO, CNPJ, celular, Cod_Desbloqueio FROM empresas_desbloueio WHERE CNPJ = '" & EA(vCnpj) & "'"
    'iRetorno = WhatsAppEnviar(SQLExecutaRetorno(sSQL, "CNPJ", ""), SQLExecutaRetorno(sSQL, "RAZAO", ""), SQLExecutaRetorno(sSQL, "celular", ""), SQLExecutaRetorno(sSQL, "Cod_Desbloqueio", ""))
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
    
'    Call ekZap.CarregarConfiguracoes
'
'    iRetorno = ekZap.setarDelay(5, mensagemRetorno, mensagemErro)
'    If Not iRetorno Then
'       MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
'       GoTo Caifora
'    Else
'
'    End If

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
    mensagemEnvio = txtMensagem.TEXT
    mensagemEnvio = Substitui(mensagemEnvio, "{cliente}", Trim(RazaoSocial), SO_UM)
    mensagemEnvio = Substitui(mensagemEnvio, "{codigodesbloqueio}", Trim(vMensagemCodigo), SO_UM)
    'ekZap.FTPurl = "ftp.ekklesiasoft.com.br"
    'ekZap.FTPuser = "onlineinfo@ekklesiasoft.com.br"
    'ekZap.FTPpassw = "Webier@online"
    'ekZap.HTTPurl = "http://www.ekklesiasoft.com.br/zapzap/onlineinfo"
    'iRetorno = ekZap.ConfigurarDLL("onlineinfo@ekklesiasoft.com.br", "6d3a55c40463afc5f5824e031bb0d0b445907", "6464", eProvedor_Solutek, mensagemErro)
    iRetorno = ekZap.configurarDLL("onlineinfo@ekklesiasoft.com.br", "26f4c100ffec076234198d8c10e85b845839b3f40340b2a212f23cf137c13f7e", "onlineinfo@ekklesiasoft.com.br", eProvedor_EkklesiaSoft, mensagemErro)

    DoEvents
    iRetorno = ekZap.enviar(idmsg, nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, mensagemRetorno, mensagemErro)
    'iRetorno = WhatsAppMensagemEnvio(nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, mensagemRetorno, mensagemErro)

    If Not iRetorno Then
       MsgBox "ERRO ENVIO: " & mensagemErro, vbCritical + vbOKOnly
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
    
    'iRetorno = ekZap.setarDelay(3, mensagemRetorno, mensagemErro)
    'If Not iRetorno Then
    '   MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
    '   GoTo Caifora
    'End If
    
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
    
    'Call ekZap.CarregarConfiguracoes
    
    'iRetorno = ekZap.setarDelay(5, mensagemRetorno, mensagemErro)
    'If Not iRetorno Then
    '   MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
    '   GoTo Caifora
    'Else
    '
    'End If

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
    'nTelefone = Retira(nTelefone, "()- ", UM_A_UM)
        
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
    'ekZap.FTPurl = "ftp.ekklesiasoft.com.br"
    'ekZap.FTPuser = "onlineinfo@ekklesiasoft.com.br"
    'ekZap.FTPpassw = "Webier@online"
    'ekZap.HTTPurl = "http://www.ekklesiasoft.com.br/zapzap/onlineinfo"
    iRetorno = ekZap.configurarDLL("onlineinfo@ekklesiasoft.com.br", "26f4c100ffec076234198d8c10e85b845839b3f40340b2a212f23cf137c13f7e", "onlineinfo@ekklesiasoft.com.br", eProvedor_EkklesiaSoft, mensagemErro)
    DoEvents
    'iRetorno = ekZap.Enviar(idmsg, nTelefone, mensagemEnvio, mensagemRetorno, mensagemErro)
    iRetorno = ekZap.enviar(idmsg, nTelefone, mensagemEnvio, mensagemRetorno, mensagemErro)
    'iRetorno = WhatsAppMensagemEnvio(nTelefone, mensagemEnvio, mensagemRetorno, mensagemErro)
'    iRetorno = ekZap.ConfigurarDLL("webieronline@gmail.com", "df925a69bcfbca2ae8465cf40e869fa8519253", "5878", mensagemErro)
  

    If Not iRetorno Then
       MsgBox "ERRO ENVIO: " & mensagemErro, vbCritical + vbOKOnly
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
    
    'iRetorno = ekZap.setarDelay(3, mensagemRetorno, mensagemErro)
    'If Not iRetorno Then
    '   MsgBox "ERRO DLL: " & mensagemErro, vbCritical + vbOKOnly
    '   GoTo Caifora
    'End If
    
    'lblStatus.Caption = "Envio de mensagem concluído com sucesso!"
    'DoEvents
    
    'ComandoSQL = "UPDATE empresas_desbloueio SET Enviado = 1 WHERE CNPJ = '" & CNPJ & "'"
    'SQLExecuta ComandoSQL
    'DoEvents
    
    Set ekZap = Nothing
    
    WhatsAppEnviarCodigo = True
    
    DoEvents
    
    Exit Function
   Resume
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
    
    Call ekZap.carregarConfiguracoes
    
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
    mensagemEnvio = txtMensagem.TEXT
    mensagemEnvio = Substitui(mensagemEnvio, "{cliente}", Trim(RazaoSocial), SO_UM)
    ekZap.FTPurl = "ftp.ekklesiasoft.com.br"
    ekZap.FTPuser = "onlineinfo@ekklesiasoft.com.br"
    ekZap.FTPpassw = "Webier@online"
    ekZap.HTTPurl = "http://www.ekklesiasoft.com.br/zapzap/onlineinfo/"
    iRetorno = ekZap.configurarDLL("webieronline@gmail.com", "cd2582694794062678efd0f9157bd299986371", "6086", eProvedor_Solutek, mensagemErro)
    DoEvents
    If Right(txtCaminho.TEXT, 3) = "jpg" Or Right(txtCaminho.TEXT, 4) = "jpeg" Then
       iRetorno = ekZap.enviarimagem(idmsg, nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, txtCaminho.TEXT, mensagemRetorno, mensagemErro)
    Else
       iRetorno = ekZap.enviararquivo(idmsg, nTelefone, Trim(RazaoSocial) & vbNewLine & mensagemEnvio, txtCaminho.TEXT, mensagemRetorno, mensagemErro)
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
    txtCaminho.TEXT = CommonDialog1.FileName
    
    If (CommonDialog1.FileName = "") Then Exit Sub
    
    'If Not FSys.FileExists(Caminho & CommonDialog1.FileTitle) Then  'se o arquivo nao existir na pasta ele copia
    '   FileCopy txtCaminho.Text, Caminho & CommonDialog1.FileTitle
    'End If
       'txtCaminho.Text = Caminho & CommonDialog1.FileTitle
    txtCaminho.TEXT = CommonDialog1.FileName
       'picLogo.Picture = LoadPicture(txtCaminho.Text) 'mostrar a imagem
    Set FSys = Nothing
End Sub

Private Sub cmdBuscarIBPT_Click()
On Error GoTo CancelaBrowse
CommonDialog1.Filter = "Arquivos CSV (*.csv)|*.csv"
CommonDialog1.FilterIndex = 1
CommonDialog1.DialogTitle = "Selecionar Tabela IBPT"
CommonDialog1.Flags = &H1000        'OFN_FILEMUSTEXIST
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
txtLocalIBPT.Text = CommonDialog1.FileName
Exit Sub
CancelaBrowse:
End Sub

Private Sub cmdEnviarIBPT_Click()
'envia manualmente a tabela IBPT (baixada por mim, desenvolvedor) pro Shared Drive "IBPT"
'no Google Drive - de la, a rotina automatica (Etapa 2, OnlineCommerce/PDV) vai buscar,
'baixar e importar a tabela em cada cliente todo dia 30
If txtLocalIBPT.Text = "" Then
    MsgBox "Selecione o arquivo da tabela IBPT antes de enviar.", vbExclamation, "Tabela IBPT"
    Exit Sub
End If

Dim FSys As FileSystemObject
Set FSys = New FileSystemObject
If Not FSys.FileExists(txtLocalIBPT.Text) Then
    MsgBox "Arquivo não encontrado: " & txtLocalIBPT.Text, vbExclamation, "Tabela IBPT"
    Set FSys = Nothing
    Exit Sub
End If
Set FSys = Nothing

Dim bEnviouIBPT As Boolean
Screen.MousePointer = 11
mensagemErro = ""
bEnviouIBPT = GoogleEnviarArquivo(txtLocalIBPT.Text, "IBPT", "0AI0VAvFMSupDUk9PVA")
Screen.MousePointer = 0

If bEnviouIBPT Then
    MsgBox "Tabela IBPT enviada para a nuvem com sucesso!", vbInformation, "Tabela IBPT"
Else
    MsgBox "Falha ao enviar a tabela IBPT." & IIf(mensagemErro <> "", vbCrLf & vbCrLf & mensagemErro, ""), vbCritical, "Tabela IBPT"
End If
End Sub

Private Sub cmdMostrarTudo_Click()
    MostrarEmpresa
End Sub

Private Sub Form_Load()
    Set moCombo = New cComboHelper
    MostrarEmpresa
    AlqIniciar
    
    If ZapConectado Then
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
    txtFantasia.TEXT = (Grid.TextMatrix(Grid.Row, 1))
    txtRazao.TEXT = (Grid.TextMatrix(Grid.Row, 2))
    mskCPF.TEXT = (Grid.TextMatrix(Grid.Row, 3))
    mskCelular.TEXT = (Grid.TextMatrix(Grid.Row, 6))
    cboCidade.TEXT = (Grid.TextMatrix(Grid.Row, 8))
    cboEstado.TEXT = (Grid.TextMatrix(Grid.Row, 9))
    If (Grid.TextMatrix(Grid.Row, 10)) = "NÃO" Then
        optCNPJ.Value = True
        optCPF.Value = False
    Else
        optCNPJ.Value = False
        optCPF.Value = True
    End If
    
    txtCodDesbloqueio.TEXT = ""
    txtCodDesbloqueioTemp.TEXT = ""
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
    Clipboard.SetText mskCPF.TEXT
End Sub

Private Sub cmdAdicionar_Click()
    If txtFantasia.TEXT = "" Or txtRazao.TEXT = "" Or mskCPF.TEXT = "" Then Exit Sub
    
    sSQL = "SELECT CNPJ FROM  empresas_desbloueio WHERE CNPJ = '" & EA(mskCPF.TEXT) & "';"
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
       EA(txtFantasia.TEXT) & "', '" & EA(txtRazao.TEXT) & "', '" & EA(mskCPF.TEXT) & "', '" & EA(mskCelular.TEXT) & "', " & vNovoCodigo & ", 0, '" & EA(cboCidade.TEXT) & "', '" & EA(cboEstado.TEXT) & "', '" & Abs(optCPF.Value) & "')"
    
    'Retorna o resultado da atualização
    Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdNovo_Click()
    LimparEmpresa
    MostrarEmpresa
End Sub

Private Sub LimparEmpresa()
    txtFantasia.TEXT = ""
    txtRazao.TEXT = ""
    cboCidade.TEXT = ""
    cboEstado.TEXT = ""
    mskCPF.Mask = ""
    mskCPF.TEXT = ""
    mskCelular.Mask = ""
    mskCelular.TEXT = ""
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
    sSQL = "SELECT *, (CASE WHEN marcado = 1 THEN 'SIM' ELSE 'NÃO' END) as vMarcado,(CASE WHEN enviado = 1 THEN 'SIM' ELSE 'NÃO' END) as vEnviado, (CASE WHEN CPF = 1 THEN 'SIM' ELSE 'NÃO' END) as vCPF  FROM  empresas_desbloueio where (FANTASIA LIKE '%" & EA(txtFantasia.TEXT) & "%')"
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

Private Sub cmdAnual_Click()
Dim vAnualAtual As Long
    i = Grid.Row
    vAnualAtual = SQLExecutaRetorno("SELECT anual FROM empresas_desbloueio WHERE CODIGO = " & Grid.TextMatrix(i, 4), "anual", 0)
    
    If vAnualAtual = 1 Then
       dbData.Execute "UPDATE empresas_desbloueio SET anual = 0 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
    Else
       dbData.Execute "UPDATE empresas_desbloueio SET anual = 1 WHERE (CODIGO = " & Grid.TextMatrix(i, 4) & ");"
    End If
    
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
    
    If txtFantasia.TEXT = "" Or txtRazao.TEXT = "" Or mskCPF.TEXT = "" Then Exit Sub
    If cboMes.TEXT = "" Then Exit Sub
    If cboAno.TEXT = "" Then Exit Sub

    vCnpj = SomarDigitos(mskCPF.TEXT)
    vQuantRazao = Len(txtRazao.TEXT)
    
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
        
    txtCodDesbloqueio.TEXT = vCodDesbloqueio
    txtCodDesbloqueioTemp.TEXT = vCodDesbTemp
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

    If txtCodDesbloqueio.TEXT = "" Then Exit Sub
    
    vUltLetra = Right(txtCodDesbloqueio.TEXT, 1)
    vMsgID = "Olá " & Trim(txtFantasia) & ", Tudo bem?"
    vMsgCode = "Seu código de desbloqueio é: " & txtCodDesbloqueio.TEXT & " referente ao mês: " & cboMes.TEXT & "/" & cboAno.TEXT & ". Use-o quando seu sistema bloquear!"
    vMsgOBS = "Obs: O último caractere do código é a letra " & vUltLetra & ". "
    
    Clipboard.Clear
    Clipboard.SetText vMsgID & vbNewLine & Trim(vMsgCode) & vbNewLine & vMsgOBS
    'Clipboard.SetText "[MENSAGEM AUTOMÁTICA]: Seu código de desbloqueio é: " & txtCodDesbloqueio.Text & " referente ao mês: Julho/2023 " & vbNewLine & " Obs: O último caractere é a letra '" & vUltLetra & "'. "
End Sub

Private Sub cmdPrepara2_Click()
    Clipboard.Clear
    Clipboard.SetText "Seu código de desbloqueio temporário é: " & txtCodDesbloqueioTemp.TEXT
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazao_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

' ============================================================
' ABA "ALIQUOTAS CBS/IBS" - regras publicadas na nuvem
' ============================================================
' A grade e a fonte das regras. "Gerar e enviar" grava TODAS as regras num arquivo texto com uma VERSAO
' (AAAAMMDDHHMMSS) e sobe para a pasta "ALIQUOTAS" do Shared Drive "IBPT". OnlineCommerce e PDV conferem a nuvem
' na abertura (1x por dia) e, se a versao for maior que a aplicada no banco do cliente, gravam as regras em
' CBS_Aliquota / IBS_Estado / IBS_Municipio. Formato do arquivo (aliquota com ponto, datas AAAAMMDD):
'   VERSAO;20261110153000
'   TIPO;ESCOPO;ALIQUOTA;INICIO;FIM;OBSERVACAO      TIPO = CBS | IBS_UF | IBS_MUN
'   ESCOPO = TODOS | UF (SP) | codigo IBGE do municipio (so IBS_MUN)
Private Function AlqPasta() As String
    Dim s As String
    s = App.Path
    If Right$(s, 1) <> "\" Then s = s & "\"
    AlqPasta = s
End Function

Private Sub AlqIniciar()
    cboAlqTipo.Clear
    cboAlqTipo.AddItem "CBS"
    cboAlqTipo.AddItem "IBS_UF"
    cboAlqTipo.AddItem "IBS_MUN"
    cboAlqTipo.ListIndex = 0
    With grdAlq
        .Rows = 1
        .Cols = 6
        .FixedCols = 0
        .ColWidth(0) = 1200: .TextMatrix(0, 0) = "Tipo"
        .ColWidth(1) = 1200: .TextMatrix(0, 1) = "Escopo"
        .ColWidth(2) = 1200: .TextMatrix(0, 2) = "Alíq. %"
        .ColWidth(3) = 1400: .TextMatrix(0, 3) = "Início"
        .ColWidth(4) = 1400: .TextMatrix(0, 4) = "Fim"
        .ColWidth(5) = 4900: .TextMatrix(0, 5) = "Observação"
        .SelectionMode = 1
        .AllowUserResizing = 1
    End With
    If Dir$(AlqPasta & "AliquotasReforma.csv") <> "" Then AlqLerArquivo AlqPasta & "AliquotasReforma.csv"
    AlqAtualizarVersao
End Sub

Private Sub AlqAtualizarVersao()
    If alqVersaoCarregada = "" Then
        lblAlqVersao.Caption = "Nenhuma versão enviada ainda a partir desta grade."
    Else
        lblAlqVersao.Caption = "Última versão gravada/carregada: " & alqVersaoCarregada & "   (" & (grdAlq.Rows - 1) & " regra(s))"
    End If
End Sub

Private Sub AlqLimparCampos()
    cboAlqTipo.ListIndex = 0
    txtAlqEscopo.Text = "TODOS"
    txtAlqAliq.Text = "": txtAlqIni.Text = "": txtAlqFim.Text = "": txtAlqObs.Text = ""
End Sub

Private Function AlqDataSQL(ByVal s As String) As String
    ' dd/mm/aaaa -> AAAAMMDD ("" se vazia)
    If Trim$(s) = "" Then AlqDataSQL = "": Exit Function
    AlqDataSQL = Format(CDate(s), "yyyymmdd")
End Function

' Le o arquivo para a grade (substitui o conteudo). Retorna "" se OK, ou o motivo.
Private Function AlqLerArquivo(ByVal sArq As String) As String
    Dim nArq As Integer, sLinha As String, a() As String
    Dim sVersao As String, nLinha As Long, r As Long
    Dim dI As Date, dF As Date
    On Error GoTo ErrLer
    grdAlq.Rows = 1
    nArq = FreeFile
    Open sArq For Input As #nArq
    Do While Not EOF(nArq)
        Line Input #nArq, sLinha
        sLinha = Trim$(sLinha)
        nLinha = nLinha + 1
        If sLinha <> "" Then
            a = Split(sLinha, ";")
            If UCase$(a(0)) = "VERSAO" Then
                If UBound(a) >= 1 Then sVersao = Trim$(a(1))
            ElseIf UBound(a) >= 4 Then
                dI = DateSerial(Val(Left$(a(3), 4)), Val(Mid$(a(3), 5, 2)), Val(Right$(a(3), 2)))
                grdAlq.AddItem UCase$(a(0)) & vbTab & UCase$(a(1)) & vbTab & Format(Val(a(2)), "0.00##") & vbTab & Format(dI, "dd/mm/yyyy") & vbTab & ""
                r = grdAlq.Rows - 1
                If Trim$(a(4)) <> "" Then
                    dF = DateSerial(Val(Left$(a(4), 4)), Val(Mid$(a(4), 5, 2)), Val(Right$(a(4), 2)))
                    grdAlq.TextMatrix(r, 4) = Format(dF, "dd/mm/yyyy")
                End If
                If UBound(a) >= 5 Then grdAlq.TextMatrix(r, 5) = a(5)
            End If
        End If
    Loop
    Close #nArq
    alqVersaoCarregada = sVersao
    AlqAtualizarVersao
    Exit Function
ErrLer:
    AlqLerArquivo = "linha " & nLinha & ": " & Err.Description
    On Error Resume Next
    Close #nArq
End Function

' Grava a grade no arquivo (cabecalho VERSAO + uma linha por regra)
Private Sub AlqGravarArquivo(ByVal sArq As String, ByVal sVersao As String)
    Dim nArq As Integer, r As Long
    Dim sAliq As String
    nArq = FreeFile
    Open sArq For Output As #nArq
    Print #nArq, "VERSAO;" & sVersao
    For r = 1 To grdAlq.Rows - 1
        sAliq = Replace(CStr(Val(Replace(grdAlq.TextMatrix(r, 2), ",", "."))), ",", ".")
        Print #nArq, grdAlq.TextMatrix(r, 0) & ";" & grdAlq.TextMatrix(r, 1) & ";" & sAliq & ";" & _
                     AlqDataSQL(grdAlq.TextMatrix(r, 3)) & ";" & AlqDataSQL(grdAlq.TextMatrix(r, 4)) & ";" & _
                     Replace(grdAlq.TextMatrix(r, 5), ";", ",")
    Next
    Close #nArq
End Sub

' Guarda a grade em AliquotasReforma.csv (rascunho local, mantem a versao ja carregada/enviada)
Private Sub AlqSalvarRascunho()
    On Error Resume Next
    AlqGravarArquivo AlqPasta & "AliquotasReforma.csv", alqVersaoCarregada
    AlqAtualizarVersao
End Sub

Private Sub grdAlq_Click()
    If grdAlq.Row < 1 Then Exit Sub
    cboAlqTipo.Text = grdAlq.TextMatrix(grdAlq.Row, 0)
    txtAlqEscopo.Text = grdAlq.TextMatrix(grdAlq.Row, 1)
    txtAlqAliq.Text = grdAlq.TextMatrix(grdAlq.Row, 2)
    txtAlqIni.Text = grdAlq.TextMatrix(grdAlq.Row, 3)
    txtAlqFim.Text = grdAlq.TextMatrix(grdAlq.Row, 4)
    txtAlqObs.Text = grdAlq.TextMatrix(grdAlq.Row, 5)
End Sub

Private Sub cmdAlqLimpar_Click()
    AlqLimparCampos
End Sub

Private Sub cmdAlqIncluir_Click()
    Dim sTipo As String, sEsc As String, sAliq As String, sIni As String, sFim As String
    Dim dAliq As Double, k As Long, r As Long, rAlvo As Long
    Dim sI2 As String, sF2 As String

    sTipo = cboAlqTipo.Text
    sEsc = UCase$(Trim$(txtAlqEscopo.Text))
    If sEsc = "" Then sEsc = "TODOS"
    sAliq = Replace(Trim$(txtAlqAliq.Text), ",", ".")
    dAliq = Val(sAliq)

    If Not (sEsc = "TODOS" Or sEsc Like "[A-Z][A-Z]" Or sEsc Like "#######") Then
        MsgBox "Escopo inválido. Use TODOS, a sigla da UF (ex.: SP) ou o código IBGE do município (7 dígitos).", vbExclamation: txtAlqEscopo.SetFocus: Exit Sub
    End If
    If sTipo = "CBS" And sEsc <> "TODOS" Then
        MsgBox "A CBS é nacional: use o escopo TODOS.", vbExclamation: txtAlqEscopo.SetFocus: Exit Sub
    End If
    If sTipo = "IBS_UF" And sEsc Like "#######" Then
        MsgBox "IBS_UF aceita TODOS ou a sigla da UF. Para município use o tipo IBS_MUN.", vbExclamation: txtAlqEscopo.SetFocus: Exit Sub
    End If
    If sAliq = "" Or dAliq < 0 Or dAliq > 100 Then
        MsgBox "Informe a alíquota em % (de 0 a 100).", vbExclamation: txtAlqAliq.SetFocus: Exit Sub
    End If
    If Not IsDate(txtAlqIni.Text) Then
        MsgBox "Informe o início da vigência (dd/mm/aaaa).", vbExclamation: txtAlqIni.SetFocus: Exit Sub
    End If
    If Trim$(txtAlqFim.Text) <> "" Then
        If Not IsDate(txtAlqFim.Text) Then
            MsgBox "Fim da vigência inválido. Use dd/mm/aaaa ou deixe vazio (sem fim).", vbExclamation: txtAlqFim.SetFocus: Exit Sub
        End If
        If CDate(txtAlqFim.Text) < CDate(txtAlqIni.Text) Then
            MsgBox "O fim da vigência não pode ser anterior ao início.", vbExclamation: txtAlqFim.SetFocus: Exit Sub
        End If
    End If
    sIni = AlqDataSQL(txtAlqIni.Text)
    sFim = AlqDataSQL(txtAlqFim.Text)
    If Trim$(txtAlqFim.Text) = "" Then sF2 = "99991231" Else sF2 = sFim

    ' mesma chave (tipo + escopo + início) = altera a linha; períodos do mesmo tipo/escopo não podem se sobrepor
    rAlvo = 0
    For r = 1 To grdAlq.Rows - 1
        If grdAlq.TextMatrix(r, 0) = sTipo And grdAlq.TextMatrix(r, 1) = sEsc Then
            If AlqDataSQL(grdAlq.TextMatrix(r, 3)) = sIni Then
                rAlvo = r
            Else
                sI2 = AlqDataSQL(grdAlq.TextMatrix(r, 3))
                If Trim$(grdAlq.TextMatrix(r, 4)) = "" Then k = 99991231 Else k = CLng(AlqDataSQL(grdAlq.TextMatrix(r, 4)))
                If CLng(sI2) <= CLng(sF2) And k >= CLng(sIni) Then
                    MsgBox "Este período se sobrepõe a outra regra do mesmo tipo e escopo (início " & grdAlq.TextMatrix(r, 3) & ")." & vbCrLf & _
                           "Ajuste o fim da regra anterior antes de incluir.", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    Next

    If rAlvo = 0 Then
        grdAlq.AddItem sTipo & vbTab & sEsc & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
        rAlvo = grdAlq.Rows - 1
    End If
    grdAlq.TextMatrix(rAlvo, 0) = sTipo
    grdAlq.TextMatrix(rAlvo, 1) = sEsc
    grdAlq.TextMatrix(rAlvo, 2) = Format(dAliq, "0.00##")
    grdAlq.TextMatrix(rAlvo, 3) = Format(CDate(txtAlqIni.Text), "dd/mm/yyyy")
    If sFim = "" Then grdAlq.TextMatrix(rAlvo, 4) = "" Else grdAlq.TextMatrix(rAlvo, 4) = Format(CDate(txtAlqFim.Text), "dd/mm/yyyy")
    grdAlq.TextMatrix(rAlvo, 5) = Replace(Trim$(txtAlqObs.Text), ";", ",")
    AlqSalvarRascunho
    AlqLimparCampos
End Sub

Private Sub cmdAlqExcluir_Click()
    If grdAlq.Row < 1 Or grdAlq.Rows < 2 Then
        MsgBox "Selecione uma regra na grade.", vbExclamation: Exit Sub
    End If
    If MsgBox("Excluir a regra " & grdAlq.TextMatrix(grdAlq.Row, 0) & " / " & grdAlq.TextMatrix(grdAlq.Row, 1) & " (início " & grdAlq.TextMatrix(grdAlq.Row, 3) & ")?" & vbCrLf & vbCrLf & _
              "Atenção: excluir aqui não apaga a regra nos clientes que já a receberam; para encerrar uma regra, preencha o Fim.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If grdAlq.Rows = 2 Then grdAlq.Rows = 1 Else grdAlq.RemoveItem grdAlq.Row
    AlqSalvarRascunho
    AlqLimparCampos
End Sub

Private Sub cmdAlqEnviar_Click()
    Dim sVersao As String, sPasta As String, sArq As String, bOK As Boolean
    If grdAlq.Rows < 2 Then
        MsgBox "Cadastre ao menos uma regra antes de enviar.", vbExclamation: Exit Sub
    End If
    If MsgBox("Publicar as " & (grdAlq.Rows - 1) & " regra(s) de alíquota para TODOS os clientes?" & vbCrLf & vbCrLf & _
              "Os sistemas vão aplicar na próxima abertura (conferem 1 vez por dia).", vbQuestion + vbYesNo, "Alíquotas CBS/IBS") = vbNo Then Exit Sub

    sVersao = Format(Now, "yyyymmddhhnnss")
    sPasta = AlqPasta & "AliquotasReforma\"
    On Error Resume Next
    MkDir sPasta
    On Error GoTo ErrEnv
    sArq = sPasta & "AliquotasReforma_" & sVersao & ".csv"
    AlqGravarArquivo sArq, sVersao

    Screen.MousePointer = 11
    mensagemErro = ""
    bOK = GoogleEnviarArquivo(sArq, "ALIQUOTAS", "0AI0VAvFMSupDUk9PVA")
    Screen.MousePointer = 0

    If bOK Then
        alqVersaoCarregada = sVersao
        AlqSalvarRascunho
        MsgBox "Alíquotas publicadas na nuvem com sucesso!" & vbCrLf & "Versão: " & sVersao, vbInformation, "Alíquotas CBS/IBS"
    Else
        MsgBox "Falha ao enviar as alíquotas." & IIf(mensagemErro <> "", vbCrLf & vbCrLf & mensagemErro, ""), vbCritical, "Alíquotas CBS/IBS"
    End If
    Exit Sub
ErrEnv:
    Screen.MousePointer = 0
    MsgBox "Erro ao gerar o arquivo: " & Err.Description, vbCritical, "Alíquotas CBS/IBS"
End Sub

Private Sub cmdAlqNuvem_Click()
    Dim sPasta As String, sNome As String, sErro As String
    If grdAlq.Rows > 1 Then
        If MsgBox("Substituir a grade atual pelo último arquivo publicado na nuvem?", vbQuestion + vbYesNo, "Alíquotas CBS/IBS") = vbNo Then Exit Sub
    End If
    sPasta = AlqPasta & "AliquotasReforma\"
    On Error Resume Next
    MkDir sPasta
    On Error GoTo 0

    Screen.MousePointer = 11
    mensagemErro = ""
    sNome = GoogleBaixarArquivo(sPasta, "ALIQUOTAS", "0AI0VAvFMSupDUk9PVA")
    Screen.MousePointer = 0
    If Trim$(sNome) = "" Then
        MsgBox "Não foi possível baixar o arquivo da nuvem." & IIf(mensagemErro <> "", vbCrLf & vbCrLf & mensagemErro, ""), vbExclamation, "Alíquotas CBS/IBS"
        Exit Sub
    End If
    sErro = AlqLerArquivo(sPasta & sNome)
    If sErro <> "" Then
        MsgBox "Arquivo baixado, mas não pôde ser lido (" & sErro & ").", vbExclamation, "Alíquotas CBS/IBS"
        Exit Sub
    End If
    AlqSalvarRascunho
    AlqLimparCampos
    MsgBox "Grade carregada da nuvem (" & (grdAlq.Rows - 1) & " regra(s)).", vbInformation, "Alíquotas CBS/IBS"
End Sub
