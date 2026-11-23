VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCampanhaZap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promoções por WhatsApp"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   16815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCampanha 
      Caption         =   "Campanha"
      Height          =   4635
      Left            =   60
      TabIndex        =   0
      Top             =   5100
      Width           =   6255
      Begin VB.TextBox txtNome 
         Height          =   315
         Left            =   1020
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Só para o histórico (o cliente não vê). Ex.: Troca de óleo de novembro"
         Top             =   270
         Width           =   5100
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Só texto"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   4
         Top             =   660
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Texto + imagem"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   5
         Top             =   660
         Width           =   1575
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Texto + vídeo"
         Height          =   255
         Index           =   2
         Left            =   4020
         TabIndex        =   6
         Top             =   660
         Width           =   1575
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Só imagem"
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   7
         Top             =   930
         Width           =   1335
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Texto + PDF"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   8
         Top             =   930
         Width           =   1575
      End
      Begin VB.TextBox txtArquivo 
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1290
         Width           =   3780
      End
      Begin ChamaleonBtn.chameleonButton cmdArquivo 
         Height          =   375
         Left            =   4860
         TabIndex        =   11
         Top             =   1260
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Escolher..."
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
         MICON           =   "frmCampanhaZap.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboTag 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1710
         Width           =   3780
      End
      Begin ChamaleonBtn.chameleonButton cmdInserirTag 
         Height          =   375
         Left            =   4860
         TabIndex        =   14
         Top             =   1680
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Inserir"
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
         MICON           =   "frmCampanhaZap.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtMensagem 
         Height          =   2175
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2100
         Width           =   6015
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblTipo 
         Caption         =   "Conteúdo:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblArquivo 
         Caption         =   "Arquivo:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblMensagem 
         Caption         =   "Mensagem:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lblDica 
         AutoSize        =   -1  'True
         Caption         =   "Use {Olá | Oi | Bom dia} para o texto variar entre os clientes (evita bloqueio)."
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   4260
         Width           =   5160
      End
   End
   Begin VB.Frame fraPublico 
      Caption         =   "Público (os filtros marcados valem juntos)"
      Height          =   4635
      Left            =   6360
      TabIndex        =   17
      Top             =   5100
      Width           =   5835
      Begin VB.CheckBox chkClientes 
         Caption         =   "Todos (somente com celular válido)"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         ToolTipText     =   "Todos os clientes ativos com celular válido (os demais filtros, menos o das parcelas vencidas, ficam desmarcados)."
         Top             =   300
         Width           =   3975
      End
      Begin VB.CheckBox chkAniver 
         Caption         =   "Aniversariantes"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cboAniver 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   630
         Width           =   1815
      End
      Begin VB.CheckBox chkComprou 
         Caption         =   "Compraram nos últimos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1020
         Width           =   2535
      End
      Begin VB.TextBox txtComprou 
         Height          =   315
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   21
         Top             =   990
         Width           =   615
      End
      Begin VB.CheckBox chkSumido 
         Caption         =   "Sem comprar há mais de"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1380
         Width           =   2535
      End
      Begin VB.TextBox txtSumido 
         Height          =   315
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1350
         Width           =   615
      End
      Begin VB.CheckBox chkProduto 
         Caption         =   "Compraram o produto"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox txtProduto 
         Height          =   315
         Left            =   2100
         MaxLength       =   60
         TabIndex        =   27
         ToolTipText     =   "Parte da descrição do produto. Com ""Compraram nos últimos X dias"" marcado, vale só para esse período."
         Top             =   1710
         Width           =   3435
      End
      Begin VB.CheckBox chkCategoria 
         Caption         =   "Da categoria"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2100
         Width           =   1935
      End
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2070
         Width           =   3435
      End
      Begin VB.CheckBox chkCidade 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2460
         Width           =   1935
      End
      Begin VB.ComboBox cboCidade 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2430
         Width           =   3435
      End
      Begin VB.CheckBox chkBairro 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2820
         Width           =   1935
      End
      Begin VB.ComboBox cboBairro 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2790
         Width           =   3435
      End
      Begin VB.CheckBox chkGastou 
         Caption         =   "Gastaram nos últimos 12 meses mais de R$"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3180
         Width           =   3435
      End
      Begin VB.TextBox txtGastou 
         Height          =   315
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   35
         Top             =   3150
         Width           =   1215
      End
      Begin VB.CheckBox chkDevedores 
         Caption         =   "Não enviar para quem tem parcela vencida"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3540
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin ChamaleonBtn.chameleonButton cmdBuscar 
         Height          =   375
         Left            =   3780
         TabIndex        =   37
         Top             =   3960
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Buscar clientes"
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
         MICON           =   "frmCampanhaZap.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblComprou 
         Caption         =   "dias"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label lblSumido 
         Caption         =   "dias"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   1380
         Width           =   600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   4680
      Left            =   60
      TabIndex        =   39
      Top             =   60
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8255
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdMarcarTodos 
      Height          =   255
      Left            =   60
      TabIndex        =   40
      Top             =   4760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Marcar todos"
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
      MICON           =   "frmCampanhaZap.frx":0054
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
      Height          =   255
      Left            =   1440
      TabIndex        =   41
      Top             =   4760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Desmarcar todos"
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
      MICON           =   "frmCampanhaZap.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   12240
      TabIndex        =   43
      Top             =   360
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox txtTeste 
      Height          =   315
      Left            =   12240
      MaxLength       =   20
      TabIndex        =   45
      Top             =   5310
      Width           =   1815
   End
   Begin ChamaleonBtn.chameleonButton cmdTeste 
      Height          =   375
      Left            =   14100
      TabIndex        =   46
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Enviar teste"
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
      MICON           =   "frmCampanhaZap.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraEnvio 
      Caption         =   "Envio"
      Height          =   2895
      Left            =   12240
      TabIndex        =   47
      Top             =   5760
      Width           =   4455
      Begin VB.CheckBox chkRodape 
         Caption         =   "Incluir no final: ""Para não receber mais promoções, responda SAIR."""
         Height          =   615
         Left            =   120
         TabIndex        =   58
         ToolTipText     =   $"frmCampanhaZap.frx":00A8
         Top             =   1260
         Width           =   3675
      End
      Begin VB.OptionButton optQuando 
         Caption         =   "Enviar agora"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   300
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optQuando 
         Caption         =   "Agendar para:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox txtLimite 
         Height          =   315
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   53
         ToolTipText     =   "Promoções por dia (a cobrança tem limite próprio e sai primeiro). Comece com pouco: número novo = mais risco de bloqueio."
         Top             =   960
         Width           =   615
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario1 
         Height          =   315
         Left            =   2640
         TabIndex        =   51
         Tag             =   "Calendario"
         ToolTipText     =   "Escolher a data no calendário"
         Top             =   600
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
         MICON           =   "frmCampanhaZap.frx":014A
         PICN            =   "frmCampanhaZap.frx":0166
         PICH            =   "frmCampanhaZap.frx":24B9
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
         Height          =   315
         Left            =   1620
         TabIndex        =   50
         ToolTipText     =   "Data do agendamento (dd/mm/aa)"
         Top             =   600
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   315
         Left            =   3000
         TabIndex        =   52
         ToolTipText     =   "Hora do agendamento (hh:mm)"
         Top             =   600
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label lblLimite 
         Caption         =   "Limite de promoções por dia:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1050
         Width           =   2295
      End
      Begin VB.Label lblPrevisao 
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   54
         Top             =   1935
         Width           =   4215
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdEnviar 
      Height          =   375
      Left            =   12240
      TabIndex        =   55
      Top             =   8700
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Criar &campanha"
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
      MICON           =   "frmCampanhaZap.frx":480C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCampanhas 
      Height          =   375
      Left            =   14520
      TabIndex        =   56
      Top             =   8700
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Acompanhar campanhas"
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
      MICON           =   "frmCampanhaZap.frx":4828
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
      Height          =   375
      Left            =   15480
      TabIndex        =   57
      Top             =   9300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      MICON           =   "frmCampanhaZap.frx":4844
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13860
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrPrevia 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   12300
      Top             =   9300
   End
   Begin VB.Image ImgMarcada 
      Height          =   195
      Left            =   7860
      Picture         =   "frmCampanhaZap.frx":4860
      Top             =   4740
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcada 
      Height          =   195
      Left            =   8160
      Picture         =   "frmCampanhaZap.frx":6C5F
      Top             =   4740
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblContador 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   12000
      TabIndex        =   38
      Top             =   4740
      Width           =   180
   End
   Begin VB.Label lblPrevia 
      Caption         =   "Como o cliente vai ver:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   42
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblTeste 
      Caption         =   "Testar no celular:"
      Height          =   255
      Left            =   12240
      TabIndex        =   44
      Top             =   5070
      Width           =   1815
   End
End
Attribute VB_Name = "frmCampanhaZap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Promoções e avisos por WhatsApp: monta a campanha (texto e/ou imagem, vídeo, PDF), escolhe o público pelos filtros,
'mostra como o cliente vai ver (WebBrowser1, balão do WhatsApp) e grava na fila. Quem envia é a tarefa do servidor
'(modCampanha.Camp_ProcessarFila), aos poucos, com limite por dia; o acompanhamento fica em frmCampanhaLista.

Private Const TITULO As String = "Promoções por WhatsApp"

Private Const C_MARCA As Integer = 0
Private Const C_CLIENTE As Integer = 1
Private Const C_CELULAR As Integer = 2
Private Const C_CIDADE As Integer = 3
Private Const C_ULTIMA As Integer = 4
Private Const C_SITUACAO As Integer = 5

Private mCli() As CampCliente
Private mMarca() As Boolean
Private mQtd As Long
Private mTags() As String
Private mIniciou As Boolean
Private mSemEvento As Boolean
Private mHtml As String

Private Sub Form_Load()
Dim desc() As String, i As Long

mIniciou = False
Camp_ListaTags mTags, desc
For i = 0 To UBound(mTags)
    cboTag.AddItem mTags(i) & "  -  " & desc(i)
Next
cboTag.ListIndex = 0
cboAniver.AddItem "do mês"
cboAniver.AddItem "de hoje"
cboAniver.ListIndex = 0
mskData.Mask = "": mskData.Text = Format$(Date + 1, "dd/mm/yy")
mskHora.Mask = "": mskHora.Text = "09:00"
txtArquivo.Text = ""
lblContador.Caption = "Marque os filtros do público (ao lado) e clique em Buscar clientes."
ReDim mCli(0): ReDim mMarca(0): mQtd = 0
MontarGrid
End Sub

Private Sub Form_Activate()
If mIniciou Then Exit Sub
mIniciou = True
If Not Camp_BancoPronto() Then
    MsgBox "O banco de dados ainda não tem as tabelas das promoções por WhatsApp (script 191)." & vbCrLf & vbCrLf & _
           "Atualize o sistema e tente de novo.", vbExclamation, TITULO
    Unload Me
    Exit Sub
End If
If Not LicTemWhatsApp() Then
    MsgBox "O envio por WhatsApp não faz parte do plano desta empresa.", vbExclamation, TITULO
    Unload Me
    Exit Sub
End If
If Not EhServidorLocal() Then
    MsgBox "As campanhas de promoção são criadas só no computador servidor (onde fica o banco de dados)." & vbCrLf & vbCrLf & _
           "O arquivo da campanha (imagem, vídeo ou PDF) fica na pasta Campanhas do sistema, no servidor, " & _
           "e é apagado ao cancelar ou excluir a campanha.", vbExclamation, TITULO
    Unload Me
    Exit Sub
End If
Me.MousePointer = vbHourglass
CarregarCombos
txtLimite.Text = Camp_LimiteDia()
AtualizarTipo
AtualizarPrevia
AtualizarPrevisao
Me.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cmdCampanhas_Click()
frmCampanhaLista.Show vbModal, Me
End Sub

'Categorias, cidades e bairros que existem de verdade nos cadastros
Private Sub CarregarCombos()
CarregarCombo cboCategoria, "SELECT DISTINCT CATEGORIA AS v FROM produtos WHERE ISNULL(CATEGORIA, '') <> '' ORDER BY 1"
CarregarCombo cboCidade, "SELECT DISTINCT Cidade AS v FROM cliente WHERE STATUS = 1 AND ISNULL(Cidade, '') <> '' ORDER BY 1"
CarregarCombo cboBairro, "SELECT DISTINCT Bairro AS v FROM cliente WHERE STATUS = 1 AND ISNULL(Bairro, '') <> '' ORDER BY 1"
End Sub

Private Sub CarregarCombo(cbo As ComboBox, ByVal sSQL As String)
Dim rs As ADODB.Recordset

cbo.Clear
On Error GoTo fim
Set rs = dbData.OpenRecordset(sSQL)
Do While Not rs.EOF
    cbo.AddItem Trim$("" & rs("v"))
    rs.MoveNext
Loop
rs.Close
fim:
Err.Clear
End Sub

'---------------------------------------------------------------------------
' Conteúdo
'---------------------------------------------------------------------------
Private Function TipoAtual() As Integer
Dim k As Integer

For k = CAMP_SO_TEXTO To CAMP_PDF
    If optTipo(k).Value Then TipoAtual = k
Next
End Function

Private Sub optTipo_Click(Index As Integer)
If Len(txtArquivo.Text) > 0 Then
    Dim sMsg As String
    If Not ArquivoValido(txtArquivo.Text, Index, sMsg) Then txtArquivo.Text = ""
End If
AtualizarTipo
AtualizarPrevia
End Sub

Private Sub AtualizarTipo()
Dim nTipo As Integer

nTipo = TipoAtual()
txtArquivo.Enabled = (nTipo <> CAMP_SO_TEXTO)
cmdArquivo.Enabled = (nTipo <> CAMP_SO_TEXTO)
lblArquivo.Enabled = (nTipo <> CAMP_SO_TEXTO)
If nTipo = CAMP_SO_TEXTO Then txtArquivo.Text = ""
txtMensagem.Enabled = (nTipo <> CAMP_SO_IMAGEM)
cboTag.Enabled = txtMensagem.Enabled
cmdInserirTag.Enabled = txtMensagem.Enabled
lblMensagem.Enabled = txtMensagem.Enabled
End Sub

'Extensão e tamanho conforme o tipo (o servidor aceita JPG/PNG, MP4 e PDF; até 15 MB)
Private Function ArquivoValido(ByVal sArq As String, ByVal nTipo As Integer, ByRef sMsg As String) As Boolean
Dim sExt As String, nTam As Long

sExt = LCase$(Mid$(sArq, InStrRev(sArq, ".") + 1))
On Error Resume Next
nTam = FileLen(sArq)
If Err.Number <> 0 Then
    Err.Clear
    sMsg = "Arquivo não encontrado: " & sArq
    Exit Function
End If
On Error GoTo 0
Select Case nTipo
    Case CAMP_IMAGEM, CAMP_SO_IMAGEM
        If sExt <> "jpg" And sExt <> "jpeg" And sExt <> "png" Then sMsg = "Escolha uma imagem JPG ou PNG.": Exit Function
        If nTam > CAMP_MAX_IMAGEM Then sMsg = "Imagem muito grande (máximo 5 MB).": Exit Function
    Case CAMP_VIDEO
        If sExt <> "mp4" Then sMsg = "Escolha um vídeo MP4.": Exit Function
        If nTam > CAMP_MAX_ARQUIVO Then sMsg = "Vídeo muito grande (máximo 15 MB).": Exit Function
    Case CAMP_PDF
        If sExt <> "pdf" Then sMsg = "Escolha um arquivo PDF.": Exit Function
        If nTam > CAMP_MAX_ARQUIVO Then sMsg = "PDF muito grande (máximo 15 MB).": Exit Function
End Select
ArquivoValido = True
End Function

Private Sub cmdArquivo_Click()
Dim nTipo As Integer, sMsg As String

nTipo = TipoAtual()
Select Case nTipo
    Case CAMP_IMAGEM, CAMP_SO_IMAGEM: CommonDialog1.Filter = "Imagens (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png"
    Case CAMP_VIDEO: CommonDialog1.Filter = "Vídeo MP4 (*.mp4)|*.mp4"
    Case CAMP_PDF: CommonDialog1.Filter = "PDF (*.pdf)|*.pdf"
    Case Else: Exit Sub
End Select
CommonDialog1.FileName = ""
CommonDialog1.DialogTitle = "Escolha o arquivo da campanha"
CommonDialog1.Flags = &H1000& Or &H4&     'cdlOFNFileMustExist Or cdlOFNHideReadOnly
CommonDialog1.CancelError = True
On Error GoTo Cancelou
CommonDialog1.ShowOpen
On Error GoTo 0
If Not ArquivoValido(CommonDialog1.FileName, nTipo, sMsg) Then
    MsgBox sMsg, vbExclamation, TITULO
    Exit Sub
End If
txtArquivo.Text = CommonDialog1.FileName
AtualizarPrevia
Exit Sub
Cancelou:
Err.Clear
End Sub

Private Sub cmdInserirTag_Click()
If cboTag.ListIndex < 0 Then Exit Sub
txtMensagem.SelText = mTags(cboTag.ListIndex)
txtMensagem.SetFocus
End Sub

Private Sub txtMensagem_Change()
ReiniciarPrevia
End Sub

Private Sub chkRodape_Click()
ReiniciarPrevia
End Sub

'---------------------------------------------------------------------------
' Prévia (balão do WhatsApp no WebBrowser1)
'---------------------------------------------------------------------------
Private Sub ReiniciarPrevia()
If Not mIniciou Then Exit Sub
tmrPrevia.Enabled = False
tmrPrevia.Enabled = True
End Sub

Private Sub tmrPrevia_Timer()
tmrPrevia.Enabled = False
AtualizarPrevia
End Sub

Private Function MensagemAtual() As String
If TipoAtual() <> CAMP_SO_IMAGEM Then MensagemAtual = txtMensagem.Text
End Function

'Cliente da linha selecionada (para a prévia e o teste) ou um nome de exemplo
Private Function NomeExemplo() As String
Dim i As Long

If grdClientes.Row >= 1 And grdClientes.Row < grdClientes.rows Then
    i = grdClientes.RowData(grdClientes.Row)
    If i >= 1 And i <= mQtd Then NomeExemplo = mCli(i).Nome
End If
If Len(NomeExemplo) = 0 Then NomeExemplo = "MARIA DA SILVA"
End Function

Private Sub AtualizarPrevia()
Dim f As Integer, sPasta As String

On Error GoTo Falha
sPasta = Environ$("TEMP")
If Len(sPasta) = 0 Then sPasta = App.path
If Right$(sPasta, 1) <> "\" Then sPasta = sPasta & "\"
mHtml = sPasta & "oc_campanha_previa.htm"
f = FreeFile
Open mHtml For Output As #f
Print #f, Camp_PreviaHtml(MensagemAtual(), NomeExemplo(), TipoAtual(), txtArquivo.Text, (chkRodape.Value = vbChecked))
Close #f
WebBrowser1.Navigate mHtml, 4   'navNoReadFromCache: o mesmo arquivo muda a cada digitação
Exit Sub
Falha:
Err.Clear
End Sub

'---------------------------------------------------------------------------
' Público
'---------------------------------------------------------------------------
Private Sub cboCategoria_Click()
If mIniciou Then chkCategoria.Value = vbChecked
End Sub

Private Sub cboCidade_Click()
If mIniciou Then chkCidade.Value = vbChecked
End Sub

Private Sub cboBairro_Click()
If mIniciou Then chkBairro.Value = vbChecked
End Sub

'"Todos (somente com celular válido)" e os filtros de público não valem juntos (o de parcela vencida vale com qualquer um)
Private Sub chkClientes_Click()
If chkClientes.Value <> vbChecked Then Exit Sub
chkAniver.Value = vbUnchecked
chkComprou.Value = vbUnchecked
chkSumido.Value = vbUnchecked
chkProduto.Value = vbUnchecked
chkCategoria.Value = vbUnchecked
chkCidade.Value = vbUnchecked
chkBairro.Value = vbUnchecked
chkGastou.Value = vbUnchecked
End Sub

Private Sub FiltroMarcado(ByVal chk As CheckBox)
If chk.Value = vbChecked Then chkClientes.Value = vbUnchecked
End Sub

Private Sub chkAniver_Click()
FiltroMarcado chkAniver
End Sub

Private Sub chkComprou_Click()
FiltroMarcado chkComprou
End Sub

Private Sub chkSumido_Click()
FiltroMarcado chkSumido
End Sub

Private Sub chkProduto_Click()
FiltroMarcado chkProduto
End Sub

Private Sub chkCategoria_Click()
FiltroMarcado chkCategoria
End Sub

Private Sub chkCidade_Click()
FiltroMarcado chkCidade
End Sub

Private Sub chkBairro_Click()
FiltroMarcado chkBairro
End Sub

Private Sub chkGastou_Click()
FiltroMarcado chkGastou
End Sub

Private Sub cboAniver_Click()
If mIniciou Then chkAniver.Value = vbChecked
End Sub

Private Sub txtComprou_Change()
If Len(txtComprou.Text) > 0 Then chkComprou.Value = vbChecked
End Sub

Private Sub txtSumido_Change()
If Len(txtSumido.Text) > 0 Then chkSumido.Value = vbChecked
End Sub

Private Sub txtProduto_Change()
If Len(Trim$(txtProduto.Text)) > 0 Then chkProduto.Value = vbChecked
End Sub

Private Sub txtGastou_Change()
If Len(txtGastou.Text) > 0 Then chkGastou.Value = vbChecked
End Sub

Private Sub txtComprou_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtSumido_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Sub txtGastou_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Or KeyAscii = 46 Then KeyAscii = 44: Exit Sub   'vírgula
KeyAscii = SoDigito(KeyAscii)
End Sub

Private Function SoDigito(ByVal KeyAscii As Integer) As Integer
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then SoDigito = KeyAscii Else SoDigito = 0
End Function

Private Sub txtLimite_Change()
AtualizarPrevisao
End Sub

'Filtros da tela; sDescricao vai para o histórico da campanha. False = filtro incompleto (com mensagem).
Private Function LerFiltro(ByRef f As CampFiltro, ByRef sDescricao As String) As Boolean
Dim s As String

If chkClientes.Value = vbChecked Then
    f.SoCelularValido = True
    s = s & "; todos com celular válido"
End If
If chkAniver.Value = vbChecked Then
    f.Aniver = cboAniver.ListIndex + 1
    s = s & "; aniversariantes " & cboAniver.Text
End If
If chkComprou.Value = vbChecked Then
    f.ComprouDias = Val(txtComprou.Text)
    If f.ComprouDias < 1 Then MsgBox "Informe em quantos dias (Compraram nos últimos X dias).", vbExclamation, TITULO: txtComprou.SetFocus: Exit Function
    s = s & "; compraram nos últimos " & f.ComprouDias & " dias"
End If
If chkSumido.Value = vbChecked Then
    f.SumidoDias = Val(txtSumido.Text)
    If f.SumidoDias < 1 Then MsgBox "Informe há quantos dias estão sem comprar.", vbExclamation, TITULO: txtSumido.SetFocus: Exit Function
    s = s & "; sem comprar há mais de " & f.SumidoDias & " dias"
End If
If chkProduto.Value = vbChecked Then
    f.Produto = Trim$(txtProduto.Text)
    If Len(f.Produto) < 2 Then MsgBox "Digite parte do nome do produto.", vbExclamation, TITULO: txtProduto.SetFocus: Exit Function
    s = s & "; produto """ & f.Produto & """"
End If
If chkCategoria.Value = vbChecked And cboCategoria.ListIndex >= 0 Then
    f.Categoria = cboCategoria.Text
    s = s & "; categoria " & f.Categoria
End If
If chkCidade.Value = vbChecked And cboCidade.ListIndex >= 0 Then
    f.Cidade = cboCidade.Text
    s = s & "; cidade " & f.Cidade
End If
If chkBairro.Value = vbChecked And cboBairro.ListIndex >= 0 Then
    f.Bairro = cboBairro.Text
    s = s & "; bairro " & f.Bairro
End If
If chkGastou.Value = vbChecked Then
    If Not IsNumeric(txtGastou.Text) Then MsgBox "Informe o valor (Gastaram mais de R$).", vbExclamation, TITULO: txtGastou.SetFocus: Exit Function
    f.GastouAcima = CCur(txtGastou.Text)
    s = s & "; gastaram mais de R$ " & Format$(f.GastouAcima, "#,##0.00") & " em 12 meses"
End If
f.ExcluirDevedores = (chkDevedores.Value = vbChecked)
If f.ExcluirDevedores Then s = s & "; sem parcela vencida"
If Len(s) = 0 Then s = "; todos os clientes ativos"   'não acontece: a busca pede um filtro
sDescricao = Mid$(s, 3)
LerFiltro = True
End Function

Private Sub cmdBuscar_Click()
Dim f As CampFiltro, sDesc As String, i As Long

If Not LerFiltro(f, sDesc) Then Exit Sub
If f.Aniver = 0 And f.ComprouDias = 0 And f.SumidoDias = 0 And Len(f.Produto) = 0 And Len(f.Categoria) = 0 And _
   Len(f.Cidade) = 0 And Len(f.Bairro) = 0 And f.GastouAcima = 0 And Not f.SoCelularValido Then
    MsgBox "Marque pelo menos um filtro do público, ou ""Todos (somente com celular válido)"".", vbInformation, TITULO
    Exit Sub
End If
If f.SoCelularValido Then
    If MsgBox("Selecionar todos os clientes ativos com celular válido?" & vbCrLf & vbCrLf & _
              "Para proteger seu número no WhatsApp, recomendamos iniciar com um envio diário reduzido.", _
              vbQuestion + vbYesNo + vbDefaultButton2, TITULO) <> vbYes Then Exit Sub
End If
Me.MousePointer = vbHourglass
mQtd = Camp_Buscar(f, mCli)
ReDim mMarca(0 To mQtd)
For i = 1 To mQtd
    mMarca(i) = (Len(mCli(i).Motivo) = 0)
Next
MontarGrid
Me.MousePointer = vbDefault
AtualizarPrevisao
End Sub

'---------------------------------------------------------------------------
' Grid
'---------------------------------------------------------------------------
Private Sub MontarGrid()
Dim i As Long, r As Long, nOk As Long, nNao As Long, nCel As Long

mSemEvento = True
With grdClientes
    .Redraw = False
    .rows = 1
    .Cols = 6
    .TextMatrix(0, C_MARCA) = "Env."
    .TextMatrix(0, C_CLIENTE) = "Cliente"
    .TextMatrix(0, C_CELULAR) = "Celular"
    .TextMatrix(0, C_CIDADE) = "Cidade"
    .TextMatrix(0, C_ULTIMA) = "Última compra"
    .TextMatrix(0, C_SITUACAO) = "Situação"
    .ColAlignment(C_MARCA) = flexAlignCenterCenter
    .ColAlignment(C_CLIENTE) = flexAlignLeftCenter
    .ColAlignment(C_CELULAR) = flexAlignLeftCenter
    .ColAlignment(C_CIDADE) = flexAlignLeftCenter
    .ColAlignment(C_ULTIMA) = flexAlignCenterCenter
    .ColAlignment(C_SITUACAO) = flexAlignLeftCenter
    For i = 1 To mQtd
        .rows = .rows + 1
        r = .rows - 1
        .RowData(r) = i
        .TextMatrix(r, C_CLIENTE) = mCli(i).Nome
        .TextMatrix(r, C_CELULAR) = mCli(i).Celular
        .TextMatrix(r, C_CIDADE) = mCli(i).Cidade
        .TextMatrix(r, C_ULTIMA) = IIf(mCli(i).UltimaCompra > 0, Format$(mCli(i).UltimaCompra, "dd/mm/yyyy"), "-")
        If Len(mCli(i).Motivo) = 0 Then
            .TextMatrix(r, C_SITUACAO) = "Vai receber"
            nOk = nOk + 1
        Else
            .TextMatrix(r, C_SITUACAO) = mCli(i).Motivo
            PintarLinha r, &H808080
            If Left$(mCli(i).Motivo, 3) = "Não" Then nNao = nNao + 1 Else nCel = nCel + 1
        End If
        MostrarMarca r
    Next
    AjustarColunas grdClientes, ",1,5,", C_MARCA
    If .rows > 1 Then .Row = 1: .Col = C_CLIENTE
    .Redraw = True
End With
mSemEvento = False
If mQtd > 0 Then
    MostrarContador
ElseIf grdClientes.Tag = "buscou" Then
    lblContador.Caption = "Nenhum cliente com esses filtros."
End If
grdClientes.Tag = "buscou"
End Sub

'Uma linha só, alinhada com a borda direita do grid (AutoSize cresce para a direita)
Private Sub lblContador_Change()
Dim nMin As Long

nMin = cmdDesmarcar.Left + cmdDesmarcar.Width + 120
lblContador.Left = grdClientes.Left + grdClientes.Width - lblContador.Width
If lblContador.Left < nMin Then lblContador.Left = nMin
End Sub

Private Sub MostrarMarca(ByVal r As Long)
Dim i As Long

i = grdClientes.RowData(r)
If i < 1 Or i > mQtd Then Exit Sub
Dim nRow As Long, nCol As Long, bSem As Boolean

'ImgMarcada = vai receber, imgDesmarcada = não vai (desmarcado ou sem poder receber)
With grdClientes
    bSem = mSemEvento
    mSemEvento = True          'mudar Row/Col aqui não é o usuário navegando
    nRow = .Row: nCol = .Col
    .TextMatrix(r, C_MARCA) = ""
    .Row = r
    .Col = C_MARCA
    If mMarca(i) And Len(mCli(i).Motivo) = 0 Then Set .CellPicture = ImgMarcada.Picture Else Set .CellPicture = imgDesmarcada.Picture
    .CellPictureAlignment = flexAlignCenterCenter
    If nRow > 0 And nRow < .rows Then .Row = nRow
    .Col = nCol
    mSemEvento = bSem
End With
End Sub

Private Sub PintarLinha(ByVal nLinha As Long, ByVal nCor As Long)
Dim c As Integer, r As Long, cc As Integer

With grdClientes
    r = .Row: cc = .Col
    .Row = nLinha
    For c = 0 To .Cols - 1
        .Col = c
        .CellForeColor = nCor
    Next
    If r > 0 And r < .rows Then .Row = r
    .Col = cc
End With
End Sub

Private Sub Marcar(ByVal nLinha As Long, ByVal bMarcar As Boolean)
Dim i As Long

If nLinha < 1 Or nLinha >= grdClientes.rows Then Exit Sub
i = grdClientes.RowData(nLinha)
If i < 1 Or i > mQtd Then Exit Sub
If Len(mCli(i).Motivo) > 0 Then bMarcar = False
mMarca(i) = bMarcar
MostrarMarca nLinha
End Sub

Private Sub grdClientes_Click()
Dim i As Long

If grdClientes.MouseCol = C_MARCA And grdClientes.MouseRow = grdClientes.Row Then
    i = grdClientes.RowData(grdClientes.Row)
    If i >= 1 And i <= mQtd Then Marcar grdClientes.Row, Not mMarca(i)
    AtualizarPrevisao
End If
End Sub

Private Sub grdClientes_KeyPress(KeyAscii As Integer)
Dim i As Long

If KeyAscii <> 32 Then Exit Sub
i = grdClientes.RowData(grdClientes.Row)
If i >= 1 And i <= mQtd Then Marcar grdClientes.Row, Not mMarca(i)
KeyAscii = 0
AtualizarPrevisao
End Sub

Private Sub grdClientes_RowColChange()
If Not mSemEvento Then ReiniciarPrevia
End Sub

Private Sub cmdMarcarTodos_Click()
MarcarTodos True
End Sub

Private Sub cmdDesmarcar_Click()
MarcarTodos False
End Sub

Private Sub MarcarTodos(ByVal bMarcar As Boolean)
Dim r As Long

grdClientes.Redraw = False
For r = 1 To grdClientes.rows - 1
    Marcar r, bMarcar
Next
grdClientes.Redraw = True
AtualizarPrevisao
End Sub

Private Sub MostrarContador()
lblContador.Caption = "Clientes exibidos: " & mQtd & ", Clientes selecionados: " & QtdMarcados() & _
                      ". Clique no quadradinho antes do nome, caso deseja remover o envio de promoção!"
End Sub

Private Function QtdMarcados() As Long
Dim i As Long

For i = 1 To mQtd
    If mMarca(i) Then QtdMarcados = QtdMarcados + 1
Next
End Function

'Larguras pelo conteúdo; as colunas de sElasticas encolhem ou crescem para caber sem rolagem horizontal
Private Sub AjustarColunas(grd As MSFlexGrid, ByVal sElasticas As String, ByVal nColFixa As Integer)
Dim c As Long, r As Long, w() As Long, wt As Long, nTotal As Long, nDisp As Long, nElast As Long, nDif As Long

On Error Resume Next
Me.FontName = grd.Font.Name
Me.FontSize = grd.Font.Size
Me.FontBold = False
ReDim w(0 To grd.Cols - 1)
For c = 0 To grd.Cols - 1
    If c = nColFixa Then
        w(c) = 450
    Else
        For r = 0 To grd.rows - 1
            wt = Me.TextWidth(grd.TextMatrix(r, c))
            If wt > w(c) Then w(c) = wt
        Next
        w(c) = w(c) + 180
    End If
    nTotal = nTotal + w(c)
    If InStr(sElasticas, "," & c & ",") > 0 Then nElast = nElast + w(c)
Next
nDisp = grd.Width - 90
If grd.rows * grd.RowHeight(0) > grd.Height - 60 Then nDisp = nDisp - 270
If nElast > 0 And nTotal <> nDisp Then
    nDif = nDisp - nTotal
    For c = 0 To grd.Cols - 1
        If InStr(sElasticas, "," & c & ",") > 0 Then
            w(c) = w(c) + CLng(CDbl(nDif) * w(c) / nElast)
            If w(c) < 600 Then w(c) = 600
        End If
    Next
End If
For c = 0 To grd.Cols - 1
    grd.ColWidth(c) = w(c)
Next
Err.Clear
End Sub

'---------------------------------------------------------------------------
' Envio
'---------------------------------------------------------------------------
Private Sub AtualizarPrevisao()
Dim n As Long, nLim As Long, nDias As Long

If Not mIniciou Then Exit Sub
If mQtd > 0 Then MostrarContador
n = QtdMarcados()
nLim = Val(txtLimite.Text)
If nLim < 1 Then nLim = CAMP_LIMITE_PADRAO
If n = 0 Then
    lblPrevisao.Caption = "Busque e marque os clientes. O servidor envia aos poucos (a cada 30 minutos, no horário e nos dias da " & _
                          "Régua de cobrança), mesmo com esta tela fechada."
Else
    nDias = (n + nLim - 1) \ nLim
    lblPrevisao.Caption = n & " cliente(s) marcado(s). Com até " & nLim & " por dia, o envio leva cerca de " & nDias & " dia(s) de envio " & _
                          "(horário e dias da Régua de cobrança). As cobranças saem primeiro e ninguém recebe duas mensagens no mesmo dia."
End If
End Sub

'Conteúdo pronto para enviar? (para o teste e para a campanha)
Private Function ConteudoValido() As Boolean
Dim nTipo As Integer, sMsg As String

nTipo = TipoAtual()
If nTipo <> CAMP_SO_IMAGEM And Len(Trim$(txtMensagem.Text)) = 0 Then
    MsgBox "Escreva a mensagem.", vbExclamation, TITULO
    txtMensagem.SetFocus
    Exit Function
End If
If nTipo <> CAMP_SO_TEXTO Then
    If Len(txtArquivo.Text) = 0 Then
        MsgBox "Escolha o arquivo (" & LCase$(Camp_TipoTexto(nTipo)) & ").", vbExclamation, TITULO
        cmdArquivo.SetFocus
        Exit Function
    End If
    If Not ArquivoValido(txtArquivo.Text, nTipo, sMsg) Then
        MsgBox sMsg, vbExclamation, TITULO
        Exit Function
    End If
End If
ConteudoValido = True
End Function

Private Sub cmdTeste_Click()
Dim sMotivo As String, sMsg As String, sArq As String

If Not CelularValido(txtTeste.Text, sMotivo) Then
    MsgBox "Celular inválido para o teste: " & sMotivo & "." & vbCrLf & vbCrLf & "Digite com o DDD, ex.: (89) 9 9999-9999.", vbExclamation, TITULO
    txtTeste.SetFocus
    Exit Sub
End If
If Not ConteudoValido() Then Exit Sub
If TipoAtual() <> CAMP_SO_TEXTO Then sArq = txtArquivo.Text
Me.MousePointer = vbHourglass
If ZapOnlineEnviar(txtTeste.Text, Camp_Texto(MensagemAtual(), NomeExemplo(), False, False, (chkRodape.Value = vbChecked)), sArq, sMsg) Then
    Me.MousePointer = vbDefault
    MsgBox "Mensagem de teste enviada (com o nome de " & Cob_PrimeiroNome(NomeExemplo()) & ").", vbInformation, TITULO
Else
    Me.MousePointer = vbDefault
    MsgBox sMsg, vbExclamation, ZapOnlineTitulo(TITULO)
End If
End Sub

'---------------------------------------------------------------------------
' Agendamento: data (dd/mm/aa, pelo calendário ou digitada) e hora (hh:mm)
'---------------------------------------------------------------------------
Private Sub cmdCalendario1_Click()
Dim varData As Variant, fCal As Calendario

Set fCal = New Calendario
fCal.Show vbModal
varData = fCal.DateSelected
Unload fCal
Set fCal = Nothing
If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub
mskData.Mask = ""
mskData.Text = Format$(varData, "dd/mm/yy")
optQuando(1).Value = True
End Sub

Private Sub mskData_GotFocus()
SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
optQuando(1).Value = True          'digitou a data = vai agendar (no GotFocus o Tab marcaria sozinho)
If mskData.Mask = "" Then mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
If mskData.Text = "__/__/__" Then
    mskData.Mask = ""
    mskData.Text = ""
End If
End Sub

Private Sub mskHora_GotFocus()
SelectControl mskHora
End Sub

Private Sub mskHora_KeyPress(KeyAscii As Integer)
optQuando(1).Value = True
If mskHora.Mask = "" Then mskHora.Mask = "##:##"
End Sub

Private Sub mskHora_LostFocus()
If mskHora.Text = "__:__" Then
    mskHora.Mask = ""
    mskHora.Text = ""
End If
End Sub

Private Sub cmdEnviar_Click()
Dim f As CampFiltro, sDesc As String, n As Long, i As Long, Cods() As Long, Cels() As String
Dim vAgenda As Variant, nLim As Long, lId As Long, sErro As String, sArq As String

If Len(Trim$(txtNome.Text)) = 0 Then
    MsgBox "Dê um nome para a campanha (só para o histórico).", vbExclamation, TITULO
    txtNome.SetFocus
    Exit Sub
End If
If Not ConteudoValido() Then Exit Sub
n = QtdMarcados()
If n = 0 Then
    MsgBox "Busque os clientes e deixe marcados (coluna Env.) os que vão receber.", vbInformation, TITULO
    Exit Sub
End If
vAgenda = Null
If optQuando(1).Value Then
    If Not IsDate(mskData.Text) Or Not IsDate(mskHora.Text) Then
        MsgBox "Informe a data (dd/mm/aa) e a hora (hh:mm) do agendamento.", vbExclamation, TITULO
        mskData.SetFocus
        Exit Sub
    End If
    vAgenda = CDate(mskData.Text) + TimeValue(mskHora.Text)
    If vAgenda <= Now Then
        MsgBox "O agendamento precisa ser para depois de agora.", vbExclamation, TITULO
        mskData.SetFocus
        Exit Sub
    End If
End If
nLim = Val(txtLimite.Text)
If nLim < 1 Then
    MsgBox "Informe o limite de promoções por dia.", vbExclamation, TITULO
    txtLimite.SetFocus
    Exit Sub
End If
LerFiltro f, sDesc

If MsgBox("Enviar a campanha """ & Trim$(txtNome.Text) & """ para " & n & " cliente(s)?" & vbCrLf & vbCrLf & _
          IIf(IsNull(vAgenda), "O servidor começa em até 30 minutos", "O servidor começa em " & Format$(vAgenda, "dd/mm/yyyy hh:nn")) & _
          " e envia até " & nLim & " por dia, no horário da Régua de cobrança." & vbCrLf & _
          "Dá para pausar ou cancelar em Acompanhar campanhas.", vbQuestion + vbYesNo, TITULO) <> vbYes Then Exit Sub

ReDim Cods(1 To n): ReDim Cels(1 To n)
n = 0
For i = 1 To mQtd
    If mMarca(i) Then
        n = n + 1
        Cods(n) = mCli(i).CodCliente
        Cels(n) = mCli(i).Celular
    End If
Next
If TipoAtual() <> CAMP_SO_TEXTO Then sArq = txtArquivo.Text
Camp_SalvarLimite nLim
Me.MousePointer = vbHourglass
lId = Camp_Salvar(Trim$(txtNome.Text), MensagemAtual(), TipoAtual(), sArq, vAgenda, sDesc, Cods, Cels, n, sErro, (chkRodape.Value = vbChecked))
Me.MousePointer = vbDefault
If lId = 0 Then
    MsgBox "Não foi possível gravar a campanha." & vbCrLf & vbCrLf & sErro, vbExclamation, TITULO
    Exit Sub
End If
MsgBox "Campanha nº " & lId & " na fila: " & n & " cliente(s)." & vbCrLf & vbCrLf & _
       "O acompanhamento fica em Acompanhar campanhas.", vbInformation, TITULO
'Limpa para a próxima (mantém os filtros)
txtNome.Text = ""
txtMensagem.Text = ""
txtArquivo.Text = ""
optTipo(CAMP_SO_TEXTO).Value = True
mQtd = 0
ReDim mCli(0): ReDim mMarca(0)
MontarGrid
lblContador.Caption = "Campanha gravada. Para outra, marque os filtros e clique em Buscar clientes."
AtualizarPrevisao
AtualizarPrevia
End Sub
