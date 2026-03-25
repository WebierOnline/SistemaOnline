VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form NFCe_Consultar 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE NFCE"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12915
   Icon            =   "Consultar_NFCe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameAguarde 
      Height          =   1935
      Left            =   2100
      TabIndex        =   66
      Top             =   4020
      Width           =   8655
      Begin VB.Label Label7 
         Caption         =   "Por favor, aguarde..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         TabIndex        =   67
         Top             =   660
         Width           =   5055
      End
   End
   Begin VB.Frame frmClassificacao 
      Caption         =   "Classificação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   31
      Top             =   960
      Width           =   2775
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox cboCriterios 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   2535
      End
      Begin VB.ComboBox cboIndice 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Situação"
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
         TabIndex        =   37
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Organização"
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
         TabIndex        =   34
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.Frame NFCe_Consultar 
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
      Height          =   2115
      Left            =   2880
      TabIndex        =   10
      Top             =   960
      Width           =   9975
      Begin VB.OptionButton optEsc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escolhendo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   20
         Top             =   900
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.OptionButton optDig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Digitado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   19
         Top             =   900
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCodPedidoCerto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox txtCodPedido 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   540
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox txtCodCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Visible         =   0   'False
         Width           =   6075
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   540
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboProduto 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Visible         =   0   'False
         Width           =   6075
      End
      Begin VB.TextBox txtCodBarra 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   2595
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   315
         Left            =   6240
         TabIndex        =   21
         Top             =   540
         Width           =   1035
         _ExtentX        =   1826
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
         MICON           =   "Consultar_NFCe.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   1500
         TabIndex        =   22
         Tag             =   "Calendario"
         Top             =   540
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
         MICON           =   "Consultar_NFCe.frx":23EE
         PICN            =   "Consultar_NFCe.frx":240A
         PICH            =   "Consultar_NFCe.frx":475D
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
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdEnviarArquivo 
         Height          =   315
         Left            =   2760
         TabIndex        =   63
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Enviar XML"
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
         MICON           =   "Consultar_NFCe.frx":6AB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdTransmitirTodas 
         Height          =   315
         Left            =   7200
         TabIndex        =   64
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Transmitir Todas"
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
         MICON           =   "Consultar_NFCe.frx":6ACC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdInutilizarTodas 
         Height          =   315
         Left            =   8580
         TabIndex        =   65
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Inutilizar Todas"
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
         MICON           =   "Consultar_NFCe.frx":6AE8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultarTodas 
         Height          =   315
         Left            =   5820
         TabIndex        =   68
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Consultar Todas"
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
         MICON           =   "Consultar_NFCe.frx":6B04
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdChave2 
         Height          =   315
         Left            =   120
         TabIndex        =   69
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Copiar Chave"
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
         MICON           =   "Consultar_NFCe.frx":6B20
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdChave 
         Height          =   375
         Left            =   8280
         TabIndex        =   70
         Top             =   1020
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Site Sefaz - Consulta"
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
         MICON           =   "Consultar_NFCe.frx":6B3C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAtualizarCliente 
         Height          =   375
         Left            =   6900
         TabIndex        =   71
         Top             =   1020
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Atualizar Cliente"
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
         MICON           =   "Consultar_NFCe.frx":6B58
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEnviarPDF 
         Height          =   315
         Left            =   1500
         TabIndex        =   72
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Enviar PDF"
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
         MICON           =   "Consultar_NFCe.frx":6B74
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodPedido 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cód. Pedido"
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
         Left            =   180
         TabIndex        =   30
         Top             =   300
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano:"
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
         TabIndex        =   28
         Top             =   300
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
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
         Left            =   180
         TabIndex        =   27
         Top             =   300
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
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
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblCodBarra 
         AutoSize        =   -1  'True
         Caption         =   "Cód. de Barra"
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
         Left            =   180
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3675
      Left            =   60
      TabIndex        =   2
      Top             =   3180
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   6482
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   12765
      TabIndex        =   0
      Top             =   0
      Width           =   12795
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTAR NFCe"
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
         Left            =   1920
         TabIndex        =   1
         Top             =   300
         Width           =   2865
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   240
         Picture         =   "Consultar_NFCe.frx":6B90
         Top             =   0
         Width           =   1500
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   375
      Left            =   5580
      TabIndex        =   3
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "Consultar_NFCe.frx":8581
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
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   8340
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "Consultar_NFCe.frx":859D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   8700
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15090
            MinWidth        =   3528
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2999
            MinWidth        =   2999
            TextSave        =   "06/03/2026"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:57"
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
   Begin ChamaleonBtn.chameleonButton cmdTransmitir 
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Transmitir"
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
      MCOL            =   0
      MPTR            =   1
      MICON           =   "Consultar_NFCe.frx":85B9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdInutilizar 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Inutilizar"
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
      MICON           =   "Consultar_NFCe.frx":85D5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdDesvincular 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Desvincular"
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
      MICON           =   "Consultar_NFCe.frx":85F1
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
      Height          =   375
      Left            =   8340
      TabIndex        =   43
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Consultar_NFCe.frx":860D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   375
      Left            =   60
      TabIndex        =   62
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Mostrar Produtos"
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
      MICON           =   "Consultar_NFCe.frx":8629
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultar 
      Height          =   375
      Left            =   2820
      TabIndex        =   5
      Top             =   8220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Consultar"
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
      MICON           =   "Consultar_NFCe.frx":8645
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "................RESUMO................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   8460
      TabIndex        =   61
      Top             =   6945
      Width           =   2790
   End
   Begin VB.Label lblQuantCancelada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9300
      TabIndex        =   60
      Top             =   7860
      Width           =   330
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelada:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8415
      TabIndex        =   59
      Top             =   7860
      Width           =   810
   End
   Begin VB.Label lblQuantInutilizada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9300
      TabIndex        =   58
      Top             =   7620
      Width           =   330
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inutilizada:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8475
      TabIndex        =   57
      Top             =   7620
      Width           =   750
   End
   Begin VB.Label lblQuantNaoEnviada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9300
      TabIndex        =   56
      Top             =   7380
      Width           =   330
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não Enviada:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8250
      TabIndex        =   55
      Top             =   7380
      Width           =   975
   End
   Begin VB.Label lblQuantEnviada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9300
      TabIndex        =   54
      Top             =   7140
      Width           =   330
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enviada:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8595
      TabIndex        =   53
      Top             =   7140
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9840
      TabIndex        =   52
      Top             =   7860
      Width           =   405
   End
   Begin VB.Label lblTotalCancelada 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00.000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10335
      TabIndex        =   51
      Top             =   7860
      Width           =   870
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9840
      TabIndex        =   50
      Top             =   7620
      Width           =   405
   End
   Begin VB.Label lblTotalInutilizada 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00.000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10335
      TabIndex        =   49
      Top             =   7620
      Width           =   870
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GERAL................................"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   48
      Top             =   6960
      Width           =   2550
   End
   Begin VB.Label lblTotalEnviada 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00.000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10335
      TabIndex        =   47
      Top             =   7140
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9840
      TabIndex        =   46
      Top             =   7140
      Width           =   405
   End
   Begin VB.Label lblTotalNaoEnviada 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00.000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10335
      TabIndex        =   45
      Top             =   7380
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9840
      TabIndex        =   44
      Top             =   7380
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1380
      TabIndex        =   41
      Top             =   7200
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   40
      Top             =   7200
      Width           =   525
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00.000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1875
      TabIndex        =   39
      Top             =   7200
      Width           =   870
   End
   Begin VB.Label lblQuant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   780
      TabIndex        =   38
      Top             =   7200
      Width           =   330
   End
   Begin VB.Label lblTotalPrazo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   60
      TabIndex        =   42
      Top             =   6900
      Width           =   12795
   End
End
Attribute VB_Name = "NFCe_Consultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Dim rEmpresa As ADODB.Recordset
Dim NFCeContingencia As Boolean
Dim IdNFProd As Long
Dim EncontroErroNFCe As Boolean

'arquivo .ini
Public oIni As Ini

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Sub Consultar_Faturas()
Dim vQParc As Integer
Dim vQNFCe As Integer
Dim vQTotal As Integer

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)
NFCeContingencia = r!ContigenciaNFCe

sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem  = " & Grid.TextMatrix(Grid.Row, 1)
Set rNFCe = dbData.OpenRecordset(sSQL)

If rNFCe.RecordCount > 0 Then
    'contar parcelas
    sSQL = "SELECT count(COD_PEDIDO) as qParc FROM parcelas WHERE COD_PEDIDO = " & Grid.TextMatrix(Grid.Row, 1)
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.EOF Then
        vQParc = r!qParc
    Else
        vQParc = 0
    End If
    
    'contar nfce_faturas
    sSQL = "SELECT count(IdNFProd) as qNFCe FROM TbNFCe_Faturas WHERE IdNFProd = " & Grid.TextMatrix(Grid.Row, 3)
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.EOF Then
        vQNFCe = r!qNFCe
    Else
        vQNFCe = 0
    End If

    If vQParc <> vQNFCe Then
        'apagar as faturas
        sSQL = "DELETE FROM TbNFCe_Faturas WHERE IdNFProd = " & Grid.TextMatrix(Grid.Row, 3)
        dbData.Execute sSQL
        'recriar as faturas
       sSQL = "INSERT INTO [TbNFCe_Faturas] " & _
              "([IdNFProd] " & _
              ",[IDParcela] " & _
              ",[TipoPgto] " & _
              ",[Vencimento] " & _
              ",[Valor] " & _
              ",[IdBandeira] " & _
              ",[CartaoNumeroAutorizacao]) " & _
              "SELECT " & rNFCe!IdNFProd & " " & _
              "     ,NUMERO " & _
              "     ,dbo.NFCeFormaPagto(FORMA_PGTO, TIPO_CARTAO) " & _
              "     ,DATA " & _
              "     ,VALOR " & _
              "     ,'01' " & _
              "     ,'' " & _
              "FROM [parcelas] " & _
              "WHERE COD_PEDIDO = " & Grid.TextMatrix(Grid.Row, 1)
       dbData.Execute sSQL
    End If
End If
End Sub

Private Sub Consultar_NFCe()
Dim codPedido As String
Dim sSQL As String, IdNFProd As Long

    Screen.MousePointer = vbHourglass
    
    codPedido = Grid.TextMatrix(Grid.Row, 3)
    
    sSQL = "SELECT IdNFProd FROM TbNFCe WHERE IdNFProd  = " & codPedido
    IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)
    
    If IdNFProd > 0 Then
       frameAguarde.Visible = True
       DoEvents
       sSQL = "SELECT NFCeChaveAcesso, NFCeProtocolo, NFCeCancelada, NFCeCanceladaProtocolo, NFCeCanceladaJustificativa FROM TbNFCe WHERE IdNFProd = " & IdNFProd
       NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")
       consultaNFCe NFeChaveAcesso, False
       
       If cStat = 100 Then
          sSQL = "UPDATE TbNFCe SET " & _
                 "NFCeEnviada = 1, " & _
                 "NFCeProtocolo = " & NFeNumeroProtocolo & ", " & _
                 "NFCeProtocoloDataHora = '" & NFeDataHora & "' " & _
                 "WHERE IdNFProd = " & IdNFProd
          vgDb.Execute sSQL
       End If
    End If
    
    Screen.MousePointer = vbDefault
    frameAguarde.Visible = False
    DoEvents
End Sub

Private Sub ConsultarCPF()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

IdNFProd = Grid.TextMatrix(Grid.Row, 3)

'consultar o cpf/cnpj do cliente
sSQL = "SELECT CPF_CNPJ, IdNFProd FROM TbNFCe WHERE IdNFProd  = " & IdNFProd
Set r = dbData.OpenRecordset(sSQL)
    
If Not r.EOF Then
    vCPF = RetirarMascaras(r!CPF_CNPJ)
Else
    vCPF = ""
End If
    
    'validar CPF
    Select Case Len(vCPF)
        Case 0
            If Len(vCPF) = 0 Then
                vCPF = Empty
            Else
                vCPF = ""
            End If
        Case 14
CNPJDigitadoErrado:
            If Validar_CNPJ(vCPF) = False Then
                    'If vNFCeConfCPF = "SIM" Then
                        If ShowMsg("Deseja inserir o CNPJ no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            vCPF = InputBox("Informe o CNPJ do cliente:", "EMISSÃO DE NFCe", "")
                            If Not Vazio(vCPF) Then
                                If Len(vCPF) = 11 Then
                                    If Validar_CPF(vCPF) = False Then
                                        MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                        GoTo CPFDigitadoErrado
                                    Else
                                        vCPF = Format(vCPF, "000\.000\.000\-00")
                                    End If
                                ElseIf Len(vCPF) = 14 Then
                                    If Validar_CNPJ(vCPF) = False Then
                                        MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                        GoTo CNPJDigitadoErrado
                                    Else
                                        vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                    End If
                                End If
                            Else
                                vCPF = ""
                            End If
                        Else
                            vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                        End If
                    'End If
            End If
        Case 11
CPFDigitadoErrado:
            If Validar_CPF(vCPF) = False Then
                    'If vNFCeConfCPF = "SIM" Then
                        If ShowMsg("Deseja inserir o CPF no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            vCPF = InputBox("Informe o CPF do cliente:", "EMISSÃO DE NFCe", "")
                            If Not Vazio(vCPF) Then
                                If Len(vCPF) = 11 Then
                                    If Validar_CPF(vCPF) = False Then
                                        MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                        GoTo CPFDigitadoErrado
                                    Else
                                        vCPF = Format(vCPF, "000\.000\.000\-00")
                                    End If
                                ElseIf Len(vCPF) = 14 Then
                                    If Validar_CNPJ(vCPF) = False Then
                                        MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                        GoTo CNPJDigitadoErrado
                                    Else
                                        vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                    End If
                                End If
                            Else
                                vCPF = ""       'se o cpf for vazio
                            End If
                        Else
                            vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                        End If
                    'End If
            End If
        Case Is < 11
            'MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
            'mskCNPJ.SetFocus
    End Select
   
If Len(vCPF) = 11 Then
    vCPF = Format(vCPF, "000\.000\.000\-00")
ElseIf Len(vCPF) = 14 Then
    vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
Else
    vCPF = ""
End If
dbData.Execute "UPDATE TbNFCe SET CPF_CNPJ = '" & vCPF & "' WHERE IdNFProd  = " & IdNFProd
End Sub

Private Sub ConsultarProdutos()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'Dim IdNFProd As Long
'Dim EncontroErroNFCe As Boolean

IdNFProd = Grid.TextMatrix(Grid.Row, 3)

'verificando os itens do pedido
sSQL = "SELECT IdNFProd, IdNFProd_Item, IDProduto, CodBarras, DescricaoProduto, CodNcm, CFOP, Bc_Icms, ICMSCST, IPICST, COFINSCST, PISCST, UN " & _
       "FROM TbNFCe_Itens " & _
       "WHERE IdNFProd = " & IdNFProd
Set rNFCeItens = dbData.OpenRecordset(sSQL)

'Dim EncontroErroNFCe As Boolean
EncontroErroNFCe = False

 For i = 1 To rNFCeItens.RecordCount
     
     'NCM..........
     If rNFCeItens!CodBarras <> "SEM GTIN" Then
         If Len(rNFCeItens!CodBarras) > 13 Or Len(rNFCeItens!CodBarras) < 8 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com EAN incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If
    
     'CFOP..........
     If rNFCeItens!CFOP <> Empty Or rNFCeItens!CFOP = "" Or rNFCeItens!CFOP = "0" Then
         If Len(rNFCeItens!CFOP) > 4 Or Len(rNFCeItens!CFOP) < 4 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com CFOP incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If
     
     'ICMS CST..........
     If rNFCeItens!icmsCST <> Empty Or rNFCeItens!icmsCST = "" Or rNFCeItens!icmsCST = "0" Then
         If Len(rNFCeItens!icmsCST) > 3 Or Len(rNFCeItens!icmsCST) < 3 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com ICMS CST incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If

     'PIS CST..........
     If rNFCeItens!pisCST <> Empty Then
         If Len(rNFCeItens!pisCST) > 2 Or Len(rNFCeItens!pisCST) < 2 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com PIS CST incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If

     'COFINS CST..........
     If rNFCeItens!cofinsCST <> Empty Then
         If Len(rNFCeItens!cofinsCST) > 2 Or Len(rNFCeItens!cofinsCST) < 2 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com COFINS CST incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If
     
     'NCM..........
     If rNFCeItens!CodNcm <> Empty Or rNFCeItens!CodNcm = "" Or rNFCeItens!CodNcm = "0" Then
         If Len(rNFCeItens!CodNcm) > 8 Or Len(rNFCeItens!CodNcm) < 8 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com NCM incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If
     'End If
     
     'UNIDADE DE MEDIDA..........
     If rNFCeItens!UN <> Empty Then
         If Len(rNFCeItens!UN) > 2 Or Len(rNFCeItens!UN) < 1 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com UNIDADE DE MEDIDA incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!DescricaoProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            GoTo Continuar
        Else
            Exit Sub
        End If
    End If
 
 rNFCeItens.MoveNext
 Next
    
Continuar:
If EncontroErroNFCe = True Then
    NFCe_Consultar_Produtos.loadPedidos IdNFProd
    NFCe_Consultar_Produtos.Show 1
End If
End Sub

Private Sub LimprarGridNFCe()
Dim sSQL As String
Dim r As ADODB.Recordset

    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set r = dbData.OpenRecordset(sSQL)
    NFCeContingencia = r!ContigenciaNFCe
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    sSQL = "SELECT Num_OS_VD_Origem, IdNFProd, NomeRazSocial, DataEmissao, Valor_NF_Prod, (CASE WHEN Inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Não Enviada' END) END) END) AS Var_Status " & _
           "FROM TbNFCe INNER JOIN pedidos ON TbNFCe.Num_OS_VD_Origem = pedidos.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO where 1 = 0"
    
    Set r = dbData.OpenRecordset(sSQL)
    
    lblQuant.Caption = r.RecordCount
    
    FormatarGrid_NFCe r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End Sub

Private Sub Mostrar_NFCe()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varCriterio As String
Dim varIndice As String
'Dim varStatus As String

If cboStatus.Text = "" Then Exit Sub

'======================================================= CRITERIOS WHERE
If cboCriterios.Text = "TODOS" Then
    varCriterio = " WHERE 1=1 "
ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
    If txtCodPedidoCerto.Text = "" Then
        varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0 "
    Else
        varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = " & txtCodPedidoCerto.Text & " "
    End If
ElseIf cboCriterios.Text = "CLIENTE" Then
    If txtCodCliente.Text = "" Then
        varCriterio = " WHERE TbNFCe.IDCliente = 0"
    Else
        varCriterio = " WHERE TbNFCe.IDCliente = " & txtCodCliente.Text & ""
    End If
ElseIf cboCriterios.Text = "DATA" Then
    If IsDate(mskData) = True Then
        varCriterio = " WHERE (DataEmissao = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103))"
    Else
        varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0"
    End If
ElseIf cboCriterios.Text = "MENSAL" Then
    If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
    
    Dim vIndMes As Integer
    Dim vWhere As String

    If cboMes.ListCount = 0 Then
        If cboMes.Text = "Janeiro" Then
            vIndMes = cboMes.ListIndex + 2
        ElseIf cboMes.Text = "Fevereiro" Then
            vIndMes = cboMes.ListIndex + 3
        ElseIf cboMes.Text = "Março" Then
            vIndMes = cboMes.ListIndex + 4
        ElseIf cboMes.Text = "Abril" Then
            vIndMes = cboMes.ListIndex + 5
        ElseIf cboMes.Text = "Maio" Then
            vIndMes = cboMes.ListIndex + 6
        ElseIf cboMes.Text = "Junho" Then
            vIndMes = cboMes.ListIndex + 7
        ElseIf cboMes.Text = "Julho" Then
            vIndMes = cboMes.ListIndex + 8
        ElseIf cboMes.Text = "Agosto" Then
            vIndMes = cboMes.ListIndex + 9
        ElseIf cboMes.Text = "Setembro" Then
            vIndMes = cboMes.ListIndex + 10
        ElseIf cboMes.Text = "Outubro" Then
            vIndMes = cboMes.ListIndex + 11
        ElseIf cboMes.Text = "Novembro" Then
            vIndMes = cboMes.ListIndex + 12
        ElseIf cboMes.Text = "Dezembro" Then
            vIndMes = cboMes.ListIndex + 13
        End If
        
        vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
    Else
        If cboMes.ListIndex = -1 Then
        If cboMes.Text = "Janeiro" Then
            vIndMes = cboMes.ListIndex + 2
        ElseIf cboMes.Text = "Fevereiro" Then
            vIndMes = cboMes.ListIndex + 3
        ElseIf cboMes.Text = "Março" Then
            vIndMes = cboMes.ListIndex + 4
        ElseIf cboMes.Text = "Abril" Then
            vIndMes = cboMes.ListIndex + 5
        ElseIf cboMes.Text = "Maio" Then
            vIndMes = cboMes.ListIndex + 6
        ElseIf cboMes.Text = "Junho" Then
            vIndMes = cboMes.ListIndex + 7
        ElseIf cboMes.Text = "Julho" Then
            vIndMes = cboMes.ListIndex + 8
        ElseIf cboMes.Text = "Agosto" Then
            vIndMes = cboMes.ListIndex + 9
        ElseIf cboMes.Text = "Setembro" Then
            vIndMes = cboMes.ListIndex + 10
        ElseIf cboMes.Text = "Outubro" Then
            vIndMes = cboMes.ListIndex + 11
        ElseIf cboMes.Text = "Novembro" Then
            vIndMes = cboMes.ListIndex + 12
        ElseIf cboMes.Text = "Dezembro" Then
            vIndMes = cboMes.ListIndex + 13
        End If
        
            vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
        Else
            vWhere = "(MONTH(DataEmissao) = " & cboMes.ListIndex + 1 & ") "
        End If
    End If
    
    varCriterio = " WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboAno & ")"
End If

'====================================================== INDICE
If cboIndice.Text = "CÓD. PEDIDO" Then
    varIndice = " TbNFCe.Num_OS_VD_Origem "
ElseIf cboIndice.Text = "CLIENTE" Then
    If cboStatus.Text = "VAZIO" Then
        varIndice = " TbNFCe.Num_OS_VD_Origem "
    Else
        varIndice = " cliente.codigo "
    End If
ElseIf cboIndice.Text = "EMISSÃO" Then
    varIndice = " TbNFCe.DataEmissao "
ElseIf cboIndice.Text = "NUM. NFCE" Then
    varIndice = " TbNFCe.IdNFProd "
Else
    varIndice = " TbNFCe.IdNFProd "
End If

'==================================================== STATUS
Dim varStatus As String
If cboStatus.Text = "TODAS" Then
    varStatus = " "
ElseIf cboStatus.Text = "ENVIADAS" Then
    varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
ElseIf cboStatus.Text = "NÃO ENVIADAS" Then
    varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
ElseIf cboStatus.Text = "CANCELADAS" Then
    varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 1 and TbNFCe.Inutilizada = 0"
ElseIf cboStatus.Text = "INUTILIADAS" Then
    varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 1"
End If

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)
NFCeContingencia = r!ContigenciaNFCe
If r.State <> 0 Then r.Close
Set r = Nothing
'============================================= CONSULTA
sSQL = "SELECT idcliente, Num_OS_VD_Origem, IdNFProd, NomeRazSocial, DataEmissao, Valor_NF_Prod, DescontoPromocional, nfcechaveacesso, (Valor_NF_Prod - DescontoPromocional) as varTotalComDesc, '' as maquina, (CASE WHEN Inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE (CASE WHEN LEFT(NFeTipoEmissao, 1) = 9 THEN 'Contingência' ELSE 'Não Enviada' END) END) END) END) AS Var_Status " & _
       "FROM TbNFCe "

sSQL = sSQL & "" & varCriterio & " " & varStatus & " ORDER BY " & varIndice

Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

lblQuant.Caption = r.RecordCount

FormatarGrid_NFCe r

If r.State <> 0 Then r.Close
Set r = Nothing


'somar as nfce =====================================================================


'======================================================= CRITERIOS WHERE
If cboCriterios.Text = "TODOS" Then
    varCriterio = " "
ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
    If txtCodPedidoCerto.Text = "" Then
        varCriterio = " AND Num_OS_VD_Origem = 0 "
    Else
        varCriterio = " AND Num_OS_VD_Origem = " & txtCodPedidoCerto.Text & " "
    End If
ElseIf cboCriterios.Text = "CLIENTE" Then
    If txtCodCliente.Text = "" Then
        varCriterio = " AND IDCliente = 0"
    Else
        varCriterio = " AND IDCliente = " & txtCodCliente.Text & ""
    End If
ElseIf cboCriterios.Text = "DATA" Then
    If IsDate(mskData) = True Then
        varCriterio = " AND (DataEmissao = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103))"
    Else
        varCriterio = " AND Num_OS_VD_Origem = 0"
    End If
ElseIf cboCriterios.Text = "MENSAL" Then
    If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
    varCriterio = " AND " & vWhere & " AND (YEAR(DataEmissao) = " & cboAno & ")"
End If

'====================================================== INDICE
If cboIndice.Text = "CÓD. PEDIDO" Then
    varIndice = " Num_OS_VD_Origem "
ElseIf cboIndice.Text = "CLIENTE" Then
    If cboStatus.Text = "VAZIO" Then
        varIndice = " Num_OS_VD_Origem "
    Else
        varIndice = " IDCliente "
    End If
ElseIf cboIndice.Text = "EMISSÃO" Then
    varIndice = " TbNFCe.DataEmissao "
Else
    varIndice = " TbNFCe.DataEmissao "
End If

'==================================================== STATUS
'Dim varStatus As String
If cboStatus.Text = "TODAS" Then
    varStatus = " "
ElseIf cboStatus.Text = "ENVIADAS" Then
    varStatus = " WHERE NFCeEnviada = 1 and NFCeCancelada = 0 and Inutilizada = 0"
ElseIf cboStatus.Text = "NÃO ENVIADAS" Then
    varStatus = " WHERE NFCeEnviada = 0 and NFCeCancelada = 0 and Inutilizada = 0"
ElseIf cboStatus.Text = "CANCELADAS" Then
    varStatus = " WHERE NFCeEnviada = 1 and NFCeCancelada = 1 and Inutilizada = 0"
ElseIf cboStatus.Text = "INUTILIADAS" Then
    varStatus = " WHERE NFCeEnviada = 0 and NFCeCancelada = 0 and Inutilizada = 1"
End If

sSQL = "SELECT TOP (1) NFCeEnviada, " & _
    "(SELECT SUM(Valor_NF_Prod - DescontoPromocional) from TbNFCe WHERE (NFCeEnviada = 1) and (NFCeCancelada = 0) and (Inutilizada = 0) " & varCriterio & ") AS varValorEnviadas, " & _
    "(SELECT SUM(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (NFCeEnviada = 0) and (Inutilizada = 0) " & varCriterio & ") AS varValorErro, " & _
    "(SELECT SUM(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (Inutilizada = 1) " & varCriterio & ") AS varValorInutilizada, " & _
    "(SELECT SUM(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (NFCeCancelada = 1) " & varCriterio & ") AS varValorCancelada, " & _
    "(SELECT COUNT(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (NFCeEnviada = 1)  and (NFCeCancelada = 0) and (Inutilizada = 0) " & varCriterio & ") AS varQuantEnviadas, " & _
    "(SELECT COUNT(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (NFCeEnviada = 0) and (Inutilizada = 0)" & varCriterio & ") AS varQuantErro, " & _
    "(SELECT COUNT(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (Inutilizada = 1) " & varCriterio & ") AS varQuantInutilizada, " & _
    "(SELECT COUNT(Valor_NF_Prod - DescontoPromocional) FROM TbNFCe WHERE (NFCeCancelada = 1)" & varCriterio & ") AS varQuantCancelada " & _
    "FROM TbNFCe as TbNFCe_1 " & _
    " " & varStatus & " " & _
    "GROUP BY NFCeEnviada"
'sSQL = sSQL & "" & varCriterio & " " & varStatus & " " & _

'Debug.Print sSQL

'sSQL = sSQL & " WHERE IDCliente = 1  " & _

Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblTotalEnviada.Caption = Format(r("varValorEnviadas"), ocMONEY)
    lblTotalNaoEnviada.Caption = Format(r("varValorErro"), ocMONEY)
    lblTotalInutilizada.Caption = Format(r("varValorInutilizada"), ocMONEY)
    lblTotalCancelada.Caption = Format(r("varValorCancelada"), ocMONEY)
    lblQuantEnviada.Caption = r("varQuantEnviadas")
    lblQuantNaoEnviada.Caption = r("varQuantErro")
    lblQuantInutilizada.Caption = r("varQuantInutilizada")
    lblQuantCancelada.Caption = r("varQuantCancelada")
Else
    lblTotalEnviada.Caption = Format(0, ocMONEY)
    lblTotalNaoEnviada.Caption = Format(0, ocMONEY)
    lblTotalInutilizada.Caption = Format(0, ocMONEY)
    lblTotalCancelada.Caption = Format(0, ocMONEY)
    lblQuantEnviada.Caption = 0
    lblQuantNaoEnviada.Caption = 0
    lblQuantInutilizada.Caption = 0
    lblQuantCancelada.Caption = 0
End If

cmdConsultar.Visible = False
cmdTransmitir.Visible = False
cmdInutilizar.Visible = False
cmdImprimir.Enabled = False
cmdCancelar.Visible = False
cmdDesvincular.Visible = False
cmdExcluir.Visible = False

If cboStatus.Text = "NÃO ENVIADAS" Then
    cmdConsultarTodas.Visible = True
    cmdTransmitirTodas.Visible = True
    cmdInutilizarTodas.Visible = True
Else
    cmdConsultarTodas.Visible = False
    cmdTransmitirTodas.Visible = False
    cmdInutilizarTodas.Visible = False
End If

Grid.SetFocus

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Public Function NUMERODATA(ByVal monthName As String) As Integer
'Select Case monthName.ToLower
'    Case Is = "JANEIRO"
'        'Return '1'
'    Case Is = "FEVEREIRO"
'        'Return 2
'    Case Is = "MARÇO"
'        'Return 3
'    Case Is = "ABRIL"
'        'Return 4
'    Case Is = "MAIO"
'        'Return 5
'    Case Is = "JUNHO"
'        'Return 6
'    Case Is = "JULHO"
'        'Return 7
'    Case Is = "AGOSTO"
'        'Return 8
'    Case Is = "SETEMBRO"
'        'Return 9
'    Case Is = "OTUBRO"
'        'Return 10
'    Case Is = "NOVENBRO"
'        'Return 11
'    Case Is = "DEZEMBRO"
'        'Return 12
'    Case Else
'        'Return 0
'End Select
End Function
Private Sub PreencherCriterios()
Dim varTexto As String
varTexto = cboCriterios.Text
cboCriterios.Clear
cboCriterios.AddItem "TODOS"
cboCriterios.AddItem "CÓD. PEDIDO"
cboCriterios.AddItem "CLIENTE"
cboCriterios.AddItem "DATA"
cboCriterios.AddItem "MENSAL"
cboCriterios.Text = varTexto
End Sub

Private Sub PreencherIndice()
Dim varTexto As String
varTexto = cboIndice.Text
cboIndice.Clear
cboIndice.AddItem "NUM. NFCE"
cboIndice.AddItem "CÓD. PEDIDO"
cboIndice.AddItem "CLIENTE"
cboIndice.AddItem "EMISSÃO"
cboIndice.Text = varTexto
End Sub

Private Sub PreencherStatus()
Dim varTexto As String
varTexto = cboStatus.Text
cboStatus.Clear
cboStatus.AddItem "TODAS"
cboStatus.AddItem "ENVIADAS"
cboStatus.AddItem "NÃO ENVIADAS"
cboStatus.AddItem "CANCELADAS"
cboStatus.AddItem "INUTILIADAS"
cboStatus.Text = varTexto
End Sub




Private Sub cboCriterios_Change()
cboCriterios_LostFocus
End Sub

Private Sub cboCriterios_Click()
cboCriterios_LostFocus
End Sub


Private Sub cboCriterios_GotFocus()
PreencherCriterios
moCombo.AttachTo cboCriterios
End Sub


Private Sub cboCriterios_LostFocus()
'If cboStatus.Text = "" Then Exit Sub

If cboCriterios.Text = "TODOS" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cmdExibir.Left = 120
    'cmdExibir_Click
ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = True
    txtCodPedido.Visible = True
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = True
    optEsc.Visible = True
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cmdExibir.Left = 3600
    'cboStatus.ListIndex = 0
    'txtCodPedido.SetFocus
ElseIf cboCriterios.Text = "CLIENTE" Then
    lblCliente.Visible = True
    cboCliente.Visible = True
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cmdExibir.Left = 6240
    'cboStatus.ListIndex = 0
    'cboCliente.SetFocus
ElseIf cboCriterios.Text = "DATA" Then
   lblCliente.Visible = False
   cboCliente.Visible = False
   'lblDesc.Visible = False
   lblCodPedido.Visible = False
   txtCodPedido.Visible = False
   'lblCodBarra.Visible = True
   txtCodCliente.Text = ""
   optDig.Visible = False
   optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = True
    mskData.Visible = True
    cmdCal1.Visible = True
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cmdExibir.Left = 1860
   'cboStatus.ListIndex = 0
'    mskData.SetFocus
ElseIf cboCriterios.Text = "MENSAL" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = True
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cmdExibir.Left = 2760
    'cboStatus.ListIndex = 0
    'cboMES.SetFocus
ElseIf cboCriterios.Text = "PRODUTO" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = True
    cboProduto.Visible = True
    lblCodBarra.Visible = False
    txtCodBarra.Visible = False
    cmdExibir.Left = 6240
    'cboProduto.SetFocus
ElseIf cboCriterios.Text = "CÓD. BARRA" Then
    lblCliente.Visible = False
    cboCliente.Visible = False
    'lblDesc.Visible = False
    lblCodPedido.Visible = False
    txtCodPedido.Visible = False
    'lblCodBarra.Visible = False
    txtCodCliente.Text = ""
    optDig.Visible = False
    optEsc.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblData.Visible = False
    mskData.Visible = False
    cmdCal1.Visible = False
    lblProduto.Visible = False
    cboProduto.Visible = False
    'cboStatus.ListIndex = 0
    lblCodBarra.Visible = True
    txtCodBarra.Visible = True
    cmdExibir.Left = 2760
    'txtCodBarra.SetFocus
End If
End Sub


Private Sub cboIndice_GotFocus()
PreencherIndice
moCombo.AttachTo cboIndice
End Sub


Private Sub cboStatus_GotFocus()
PreencherStatus
moCombo.AttachTo cboStatus
End Sub


Private Sub cboStatus_LostFocus()
'cmdConsultarTodas
'cmdTransmitirTodas
'cmdInutilizarTodas
End Sub



Private Sub EnviaEmailPDF(EmailPara As String, Anexo1 As String)
Dim emailDest As String, pathAnexo() As String, NomeRemetente As String, corpoEmail As String, emailCC() As String
Dim temParcelas As Boolean
Dim sistNFe As snfe.Util

On Error GoTo deuErro

Set sistNFe = New snfe.Util

emailDest = EmailPara

ReDim emailCC(0)
emailCC(0) = EmailPara

If Vazio(emailDest) Then Exit Sub

iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)

ReDim pathAnexo(0)
pathAnexo(0) = Anexo1

NomeRemetente = SQLExecutaRetorno("SELECT Fantasia FROM empresa", "Fantasia", rEmpresa!Fantasia)
corpoEmail = "Segue em anexo o arquivo PDF da NFCe emitida. " & _
             "<br><br>" & _
             "Atenciosamente, " & _
             "<br><br>" & _
             "#nome_emitente#"
corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO"), SO_UM)

If (emailDest <> Empty) Then
   Screen.MousePointer = vbHourglass
   iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivo PDF referente a NFCe emitida ", corpoEmail, pathAnexo, emailCC)
   Screen.MousePointer = vbDefault
End If

If iRetorno Then MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"

Set sistNFe = Nothing

Exit Sub
    
deuErro:
    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: Envio Email"
    Err.Clear
    Set sistNFe = Nothing
End Sub
Private Sub EnviaEmail(EmailPara As String, Anexo1 As String)
Dim emailDest As String, pathAnexo() As String, NomeRemetente As String, corpoEmail As String, emailCC() As String
Dim temParcelas As Boolean
Dim sistNFe As snfe.Util

On Error GoTo deuErro

Set sistNFe = New snfe.Util

emailDest = EmailPara

ReDim emailCC(0)
emailCC(0) = EmailPara

If Vazio(emailDest) Then Exit Sub

iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)

ReDim pathAnexo(0)
pathAnexo(0) = Anexo1

NomeRemetente = SQLExecutaRetorno("SELECT Fantasia FROM empresa", "Fantasia", rEmpresa!Fantasia)
corpoEmail = "Segue em anexo o arquivo XML da NFCe emitida. " & _
             "<br><br>" & _
             "Atenciosamente, " & _
             "<br><br>" & _
             "#nome_emitente#"
corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO"), SO_UM)

If (emailDest <> Empty) Then
   Screen.MousePointer = vbHourglass
   iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivo XML referente a NFCe emitida ", corpoEmail, pathAnexo, emailCC)
   Screen.MousePointer = vbDefault
End If

If iRetorno Then MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"

Set sistNFe = Nothing

Exit Sub
    
deuErro:
    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: Envio Email"
    Err.Clear
    Set sistNFe = Nothing
End Sub


'Continuar:
'MsgBox "Chegou no final"
Private Sub cmdAtualizarCliente_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim varCodCliente As String
varCodCliente = Grid.TextMatrix(Grid.Row, 11)

If ShowMsg("Deseja atualizar o cliente " & Grid.TextMatrix(Grid.Row, 4) & " ?", vbInformation + vbYesNo) = vbYes Then

Load Clientes_Cadastro
Clientes_Cadastro.SSTab1.Tab = 0
Clientes_Cadastro.cmdNovo.Enabled = False
Clientes_Cadastro.cmdSalvar.Enabled = False
Clientes_Cadastro.cmdCancelar.Enabled = False
Clientes_Cadastro.cmdAlterar.Enabled = True
Clientes_Cadastro.cmdExcluir.Enabled = True
Clientes_Cadastro.txtCodigo.Text = varCodCliente
Clientes_Cadastro.Show 1

End If
End Sub

Private Sub cmdCancelar_Click()
If MsgBox("Tem certeza que deseja cancelar " & Chr(13) & "o cupom fiscal (NFCe): '" & Grid.TextMatrix(Grid.Row, 3) & "' ?", vbQuestion + vbYesNo, "Cancelamento") = vbYes Then
    Cancelar_NFCe Grid.TextMatrix(Grid.Row, 3)
    Mostrar_NFCe
End If
End Sub

Private Sub Cancelar_NFCe(codPedido As String)
Dim sSQL As String, IdNFProd As Long

    Screen.MousePointer = vbHourglass
    
    sSQL = "SELECT IdNFProd FROM TbNFCe WHERE IdNFProd  = " & codPedido
    IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)
    
    If IdNFProd > 0 Then
       frameAguarde.Visible = True
       DoEvents
       sSQL = "SELECT NFCeChaveAcesso, NFCeProtocolo, NFCeCancelada, NFCeCanceladaProtocolo, NFCeCanceladaJustificativa FROM TbNFCe WHERE IdNFProd = " & IdNFProd
       NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")
       NFeNumeroProtocolo = SQLExecutaRetorno(sSQL, "NFCeProtocolo", 0)
       If Not Vazio(NFeChaveAcesso) And NFeNumeroProtocolo > 0 Then
          If CancelaNFCe(NFeChaveAcesso, NFeNumeroProtocolo, "DESISTENCIA DE COMPRA", True) Then
             sSQL = "UPDATE TbNFCe SET NFCeCancelada = 1, NFceCanceladaProtocolo = " & NFeNumeroProtocolo & ", NFCeCanceladaJustificativa = 'DESISTENCIA DE COMPRA', Num_OS_VD_Origem = 0 WHERE IDNFProd = " & IdNFProd
             dbData.Execute sSQL
          End If
       End If
    End If
    
    Screen.MousePointer = vbDefault
    frameAguarde.Visible = False
    DoEvents
End Sub


Private Sub cmdChave_Click()
Dim vCHAVE As String

vCHAVE = Grid.TextMatrix(Grid.Row, 10)

Clipboard.Clear
Clipboard.SetText vCHAVE

ShellExecute hwnd, "open", "https://dfe-portal.svrs.rs.gov.br/NFCE/Consulta", vbNullString, vbNullString, conSwNo
ShellExecute hwnd, "open", "https://www.sefaz.pi.gov.br/nfce/consulta/", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdChave2_Click()
Dim vCHAVE As String

vCHAVE = Grid.TextMatrix(Grid.Row, 10)

Clipboard.Clear
Clipboard.SetText vCHAVE
End Sub


Private Sub cmdConsultar_Click()
ConsultarCPF
ConsultarProdutos
If EncontroErroNFCe = False Then
    Consultar_NFCe
    Mostrar_NFCe
End If
End Sub

Private Sub cmdConsultarTodas_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim varCriterio As String
Dim varIndice As String
Dim codPedido As String
Dim IdNFProd As Long

    Screen.MousePointer = vbHourglass

    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set r = dbData.OpenRecordset(sSQL)
    NFCeContingencia = r!ContigenciaNFCe
    
    If NFCeContingencia Then
       MsgBox "CONTINGÊNCIA DA NFCE ATIVADA, ENVIO NÃO PERMITIDO!", vbInformation + vbOKOnly
       GoTo Caifora
    End If
    
    If cboStatus.Text = "" Then Exit Sub
    
    '======================================================= CRITERIOS WHERE
    If cboCriterios.Text = "TODOS" Then
        varCriterio = " WHERE 1=1 "
    ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
        If txtCodPedidoCerto.Text = "" Then
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0 "
        Else
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = " & txtCodPedidoCerto.Text & " "
        End If
    ElseIf cboCriterios.Text = "CLIENTE" Then
        If txtCodCliente.Text = "" Then
            varCriterio = " WHERE TbNFCe.IDCliente = 0"
        Else
            varCriterio = " WHERE TbNFCe.IDCliente = " & txtCodCliente.Text & ""
        End If
    ElseIf cboCriterios.Text = "DATA" Then
        If IsDate(mskData) = True Then
            varCriterio = " WHERE (DataEmissao = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103))"
        Else
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0"
        End If
    ElseIf cboCriterios.Text = "MENSAL" Then
        If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
        
        Dim vIndMes As Integer
        Dim vWhere As String
    
        If cboMes.ListCount = 0 Then
            If cboMes.Text = "Janeiro" Then
                vIndMes = cboMes.ListIndex + 2
            ElseIf cboMes.Text = "Fevereiro" Then
                vIndMes = cboMes.ListIndex + 3
            ElseIf cboMes.Text = "Março" Then
                vIndMes = cboMes.ListIndex + 4
            ElseIf cboMes.Text = "Abril" Then
                vIndMes = cboMes.ListIndex + 5
            ElseIf cboMes.Text = "Maio" Then
                vIndMes = cboMes.ListIndex + 6
            ElseIf cboMes.Text = "Junho" Then
                vIndMes = cboMes.ListIndex + 7
            ElseIf cboMes.Text = "Julho" Then
                vIndMes = cboMes.ListIndex + 8
            ElseIf cboMes.Text = "Agosto" Then
                vIndMes = cboMes.ListIndex + 9
            ElseIf cboMes.Text = "Setembro" Then
                vIndMes = cboMes.ListIndex + 10
            ElseIf cboMes.Text = "Outubro" Then
                vIndMes = cboMes.ListIndex + 11
            ElseIf cboMes.Text = "Novembro" Then
                vIndMes = cboMes.ListIndex + 12
            ElseIf cboMes.Text = "Dezembro" Then
                vIndMes = cboMes.ListIndex + 13
            End If
            
            vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
        Else
            If cboMes.ListIndex = -1 Then
            If cboMes.Text = "Janeiro" Then
                vIndMes = cboMes.ListIndex + 2
            ElseIf cboMes.Text = "Fevereiro" Then
                vIndMes = cboMes.ListIndex + 3
            ElseIf cboMes.Text = "Março" Then
                vIndMes = cboMes.ListIndex + 4
            ElseIf cboMes.Text = "Abril" Then
                vIndMes = cboMes.ListIndex + 5
            ElseIf cboMes.Text = "Maio" Then
                vIndMes = cboMes.ListIndex + 6
            ElseIf cboMes.Text = "Junho" Then
                vIndMes = cboMes.ListIndex + 7
            ElseIf cboMes.Text = "Julho" Then
                vIndMes = cboMes.ListIndex + 8
            ElseIf cboMes.Text = "Agosto" Then
                vIndMes = cboMes.ListIndex + 9
            ElseIf cboMes.Text = "Setembro" Then
                vIndMes = cboMes.ListIndex + 10
            ElseIf cboMes.Text = "Outubro" Then
                vIndMes = cboMes.ListIndex + 11
            ElseIf cboMes.Text = "Novembro" Then
                vIndMes = cboMes.ListIndex + 12
            ElseIf cboMes.Text = "Dezembro" Then
                vIndMes = cboMes.ListIndex + 13
            End If
            
                vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
            Else
                vWhere = "(MONTH(DataEmissao) = " & cboMes.ListIndex + 1 & ") "
            End If
        End If
        
        varCriterio = " WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboAno & ")"
    End If
    
    '====================================================== INDICE
    If cboIndice.Text = "CÓD. PEDIDO" Then
        varIndice = " TbNFCe.Num_OS_VD_Origem "
    ElseIf cboIndice.Text = "CLIENTE" Then
        If cboStatus.Text = "VAZIO" Then
            varIndice = " TbNFCe.Num_OS_VD_Origem "
        Else
            varIndice = " cliente.codigo "
        End If
    ElseIf cboIndice.Text = "EMISSÃO" Then
        varIndice = " TbNFCe.DataEmissao "
    ElseIf cboIndice.Text = "NUM. NFCE" Then
        varIndice = " TbNFCe.IdNFProd "
    Else
        varIndice = " TbNFCe.IdNFProd "
    End If
    
    '==================================================== STATUS
    Dim varStatus As String
    If cboStatus.Text = "TODAS" Then
        varStatus = " "
    ElseIf cboStatus.Text = "ENVIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "NÃO ENVIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "CANCELADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 1 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "INUTILIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 1"
    End If
    
    sSQL = "SELECT IdNFProd, NFCeChaveAcesso " & _
           "FROM TbNFCe "
    sSQL = sSQL & "" & varCriterio & " " & varStatus & " ORDER BY " & varIndice
        
    Set r = dbData.OpenRecordset(sSQL)
    
    frameAguarde.Visible = True
    DoEvents
    Do While Not r.EOF
       consultaNFCe r!NFCeChaveAcesso, True
       
       If cStat = 100 Then
          sSQL = "UPDATE TbNFCe SET " & _
                 "NFCeEnviada = 1, " & _
                 "NFCeProtocolo = " & NFeNumeroProtocolo & ", " & _
                 "NFCeProtocoloDataHora = '" & NFeDataHora & "' " & _
                 "WHERE IdNFProd = " & r!IdNFProd
          vgDb.Execute sSQL
       End If
       r.MoveNext
    Loop
    
Caifora:
    Set r = Nothing
    frameAguarde.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    Mostrar_NFCe
End Sub

Private Sub cmdDesvincular_Click()
If ShowMsg("Tem certeza que deseja desvincular esse NFCe " & Grid.TextMatrix(Grid.Row, 3) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "UPDATE TbNFCe SET  Num_OS_VD_Origem = '' WHERE (IdNFProd = " & Grid.TextMatrix(Grid.Row, 3) & ");"
End If
'
Mostrar_NFCe
End Sub

Private Sub cmdEnviarArquivo_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String

'parte de encontrar o arquivo
sSQL = "SELECT DiretorioXML, fantasia FROM Empresa"
Set rEmpresa = dbData.OpenRecordset(sSQL)

If Not rEmpresa.EOF Then
    dirXML = IIf(Right(rEmpresa!DiretorioXML, 1) = "\", rEmpresa!DiretorioXML, rEmpresa!DiretorioXML & "\")
End If

IdNFProd = Val(Grid.TextMatrix(Grid.Row, 3))

sSQL = "SELECT NFCeChaveAcesso, DataEmissao FROM TbNFCe WHERE IdNFProd = " & IdNFProd
NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")
NFeDataHoraEnvio = SQLExecutaRetorno(sSQL, "DataEmissao", "")

xCaminhoXML = dirXML & "nfe\arquivos\procNFe\" & NFeChaveAcesso & "-procNFe.xml"
anoEmes = dirXML & "nfe\arquivos\procNFe\" & Format(NFeDataHoraEnvio, "yyyymm") & "\"

If Not Existe(anoEmes) Then MsgBox "Não existe a pasta referente ao mês selecionado!", vbInformation, "Aviso do Sistema": Exit Sub

If Not Existe(xCaminhoXML) Then xCaminhoXML = anoEmes & NFeChaveAcesso & "-procNFe.xml"
'verifica se o arquivo existe
If Not Existe(xCaminhoXML) Then MsgBox "Não existe o arquivo XML dessa venda nesse computador!", vbInformation, "Aviso do Sistema": Exit Sub

'envio do arquivo
emailDestino = InputBox("Informe o e-mail do destinatário", "Envio de Email", "")

If Not Vazio(emailDestino) Then
   Call EnviaEmail(emailDestino, xCaminhoXML)
   DoEvents
End If
End Sub

Private Sub cmdEnviarPDF_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String

'parte de encontrar o arquivo
sSQL = "SELECT DiretorioXML, fantasia FROM Empresa"
Set rEmpresa = dbData.OpenRecordset(sSQL)

If Not rEmpresa.EOF Then
    dirXML = IIf(Right(rEmpresa!DiretorioXML, 1) = "\", rEmpresa!DiretorioXML, rEmpresa!DiretorioXML & "\")
End If

IdNFProd = Val(Grid.TextMatrix(Grid.Row, 3))

sSQL = "SELECT NFCeChaveAcesso, DataEmissao FROM TbNFCe WHERE IdNFProd = " & IdNFProd
NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")

xCaminhoXML = dirXML & "nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"

'verifica se o arquivo existe
If Not Existe(xCaminhoXML) Then MsgBox "Não existe o arquivo PDF dessa venda nesse computador!", vbInformation, "Aviso do Sistema": Exit Sub

'envio do arquivo
emailDestino = InputBox("Informe o e-mail do destinatário", "Envio de Email", "")

If Not Vazio(emailDestino) Then
   Call EnviaEmailPDF(emailDestino, xCaminhoXML)
   DoEvents
End If

End Sub


Private Sub cmdExcluir_Click()
Consultar_NFCe
Mostrar_NFCe

Dim sSQL As String
Dim varCodPedido As Long

If Grid.TextMatrix(Grid.Row, 8) = "Enviada" Or Grid.TextMatrix(Grid.Row, 8) = "Cancelada" Or Grid.TextMatrix(Grid.Row, 8) = "Inutilizada" Then
    MsgBox "Não é possivel excluir essa NFCe!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem  = " & codPedido
'varCodPedido = SQLExecutaRetorno(sSQL, "IdNFProd", 0)

If ShowMsg("Tem certeza que deseja desvincular esse NFCe " & Grid.TextMatrix(Grid.Row, 3) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    dbData.Execute "DELETE FROM TbNFCe_Itens WHERE (IdNFProd = " & Grid.TextMatrix(Grid.Row, 3) & ");"
    dbData.Execute "DELETE FROM TbNFCe WHERE (IdNFProd = " & Grid.TextMatrix(Grid.Row, 3) & ");"
End If

Mostrar_NFCe
End Sub

Private Sub cmdExibir_Click()
    Mostrar_NFCe
End Sub
Function EImpar(ByVal iNum As Long) As Boolean
   EImpar = (iNum Mod 2)
End Function
Sub FlexCores(lCorPar As Long, lCorImpar As Long)
   'ZEBRAR O FLEXGRID
   Dim iLinha As Integer
   Dim lCor As OLE_COLOR
   
   Grid.FillStyle = flexFillRepeat
   
   For iLinha = 1 To Grid.Rows - 1
      With Grid
         .Row = iLinha
         
         If EImpar(iLinha) Then 'Se a linha for impar:
            lCor = lCorImpar
         Else
            lCor = lCorPar
         End If
         
         .Col = 1                'Seleciona a partir da primeira coluna
         .ColSel = .Cols - 1     'Seleciona até a última coluna
         .CellBackColor = lCor   'Aplica a cor
      End With
   Next
   
   Grid.FillStyle = flexFillSingle
End Sub

Private Sub SomaNFCe()
On Error GoTo errorhandeler
Dim soma As Currency
Dim i As Integer

soma = 0
With Grid
   For i = 1 To .Rows - 1
         soma = soma + CCur(.TextMatrix(i, 7))
   Next
End With

lblTotal.Caption = Format(soma, ocMONEY)
   
errorhandeler:
End Sub

Private Sub cmdExibirProdutos_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

NFCe_Consultar_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 3)
NFCe_Consultar_Produtos.Show 1
End Sub


Private Sub FormatarGrid_NFCe(rTabela As ADODB.Recordset)
   Dim i As Integer
   With Grid
      .Clear
      .Cols = 12
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 850
      .ColWidth(2) = 900
      .ColWidth(3) = 650
      .ColWidth(4) = 4000
      .ColWidth(5) = 1100
      .ColWidth(6) = 900
      .ColWidth(7) = 1100
      .ColWidth(8) = 1100
      .ColWidth(9) = 0
      .ColWidth(10) = 0
      .ColWidth(11) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "EMISSÃO"
      .TextMatrix(0, 3) = "NFCe"
      .TextMatrix(0, 4) = "CLIENTE"
      .TextMatrix(0, 5) = "SUBTOTAL"
      .TextMatrix(0, 6) = "DESC"
      .TextMatrix(0, 7) = "TOTAL"
      .TextMatrix(0, 8) = "SITUAÇÃO"
      .TextMatrix(0, 9) = "MAQ."
      .TextMatrix(0, 10) = "CHAVE"
      .TextMatrix(0, 11) = "ID"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next i
      'Num_OS_VD_Origem, IdNFProd, NomeRazSocial, DataEmissao, Valor_NF_Prod

      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("Num_OS_VD_Origem"), "000000")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DataEmissao"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("IdNFProd"), "00000")
            .TextMatrix(.Rows - 1, 4) = rTabela("NomeRazSocial")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("Valor_NF_Prod"), ocMONEY)
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("DescontoPromocional"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("varTotalComDesc"), ocMONEY)
            .TextMatrix(.Rows - 1, 8) = rTabela("var_status")
            .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("maquina"))
            .TextMatrix(.Rows - 1, 10) = ValidateNull(rTabela("nfcechaveacesso"))
            .TextMatrix(.Rows - 1, 11) = ValidateNull(rTabela("IDCLIENTE"))
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
         FlexCores &HFFFFFF, &HE0E0E0

      .Rows = .Rows - 1
      SomaNFCe
   End With
End Sub

Private Sub cmdGerarArquivo_Click()

End Sub

Private Sub cmdImprimir_Click()

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)
NFCeContingencia = r!ContigenciaNFCe

'descobrir a maquina de origem do nfce
sSQL = "SELECT DISTINCT MAQUINA FROM pedidos WHERE cod_pedido = " & Grid.TextMatrix(Grid.Row, 1) & ""
Set r = dbData.OpenRecordset(sSQL)
Debug.Print sSQL
Dim vMaquinas As String
If Not r.EOF Then
   vMaquinas = ValidateNull(r("MAQUINA"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing

'MsgBox vMaquinas

If vMaquinas <> StatusBar1.Panels(2).Text Then MsgBox "A Impressão da NFCe terá que ser realizada no PDV de origem! " & vMaquinas & " ", vbExclamation, "Aviso do Sistema": Exit Sub


Dim anoEmes As String, Arquivo As String
Dim codPedido As String, NomeImpNFCe As String
Dim IdNFProd As Long

'Set cCfg = sysConfig("NOME_IMP_NFCE")
'NomeImpNFCe = cCfg.Value
'Set cCfg = Nothing

'pegar o nome da impressora no ini
Dim oIni As Ini
''Dim var_ImpTermica As String

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
NomeImpNFCe = oIni.LerTexto("IMPRESSORA_NFCE", "impressora")
Set oIni = Nothing

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = NomeImpNFCe Then
      Set Printer = Prt
      Exit For
   End If
Next

Dim vDataEmissao As Date
vDataEmissao = Grid.TextMatrix(Grid.Row, 2)

Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

codPedido = Grid.TextMatrix(Grid.Row, 3)

sSQL = "SELECT IdNFProd FROM TbNFCe WHERE IdNFProd  = " & codPedido
IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)

If IdNFProd > 0 Then
   sSQL = "SELECT NFCeChaveAcesso, NFCeProtocolo, NFeTipoEmissao, NFCeEnviada, NFCeCancelada, NFCeCanceladaProtocolo, NFCeCanceladaJustificativa FROM TbNFCe WHERE IdNFProd = " & IdNFProd
   NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")

   dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM Empresa", "DiretorioXML", App.Path)
   dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
   If Left(SQLExecutaRetorno(sSQL, "NFeTipoEmissao", 1), 1) = 1 Or SQLExecutaRetorno(sSQL, "NFCeEnviada", False) Then
      xCaminhoXML = dirXML & "nfe\arquivos\procNFe\" & NFeChaveAcesso & "-procNFe.xml"
      anoEmes = dirXML & "nfe\arquivos\procNFe\" & Format(vDataEmissao, "yyyymm") & "\"
      If Not Existe(xCaminhoXML) Then xCaminhoXML = anoEmes & NFeChaveAcesso & "-procNFe.xml"
   ElseIf Left(SQLExecutaRetorno(sSQL, "NFeTipoEmissao", 1), 1) = 9 Or SQLExecutaRetorno(sSQL, "NFCeEnviada", False) = False Then
      xCaminhoXML = dirXML & "nfe\arquivos\assinado\NFe" & NFeChaveAcesso & "-assinado.xml"
   End If
   xCaminhoPDF = dirXML & "nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"

   If Not NFCeContingencia Then
      Call sistNFe.DANFCeImprimir(xCaminhoXML, True, NomeImpNFCe, True, xCaminhoPDF, 0, SQLExecutaRetorno(sSQL, "NFCeCancelada", False), False, "")
   ElseIf Left(SQLExecutaRetorno(sSQL, "NFeTipoEmissao", 1), 1) = 9 And SQLExecutaRetorno(sSQL, "NFCeEnviada", False) = False Then
      Call sistNFe.DANFCeOFFImprimir(xCaminhoXML, True, NomeImpNFCe, True, xCaminhoPDF, 0, SQLExecutaRetorno(sSQL, "NFCeCancelada", False), False, "")
   Else
      Call sistNFe.DANFCeImprimir(xCaminhoXML, True, NomeImpNFCe, True, xCaminhoPDF, 0, SQLExecutaRetorno(sSQL, "NFCeCancelada", False), False, "")
   End If
   'Call sistNFe.ImpNFCe(xCaminhoXML, "", "", False, NomeImpNFCe, True, xCaminhoPDF, False, False, 0, "")
   
End If
End Sub

Private Sub cmdInutilizar_Click()
Dim codPedido As String, nNota As String, CNPJ As String
Dim sSQL As String, IdNFProd As Long

codPedido = Grid.TextMatrix(Grid.Row, 3)

If MsgBox("Tem certeza que deseja inutilizar " & Chr(13) & "o cupom fiscal (NFCe): '" & codPedido & "' ?", vbQuestion + vbYesNo, "Inutilização") = vbYes Then

    dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM Empresa", "DiretorioXML", App.Path)
    dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
    CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
    
    sSQL = "SELECT IdNFProd FROM TbNFCe WHERE IdNFProd  = " & codPedido
    IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)
    If IdNFProd > 0 Then
       frameAguarde.Visible = True
       DoEvents
       sSQL = "SELECT NumeNota FROM TbNFCe WHERE IdNFProd = " & IdNFProd
       nNota = SQLExecutaRetorno(sSQL, "NumeNota", "0")
       Dim sistNFe As snfe.Util
       Set sistNFe = New snfe.Util
       iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
       iRetorno = sistNFe.InutilizarNumeracao(Format(Date, "yyyy"), CNPJ, "ERRO AO TRANSMITIR NOTA, PERDA DE SEQUENCIA", nNota, nNota, 1, xCaminhoXML)
       'NFeResposta = sistNFe.NfceInutilizacao(Format(Date, "yy"), "1", nNota, nNota, "ERRO AO TRANSMITIR NOTA, PERDA DE SEQUENCIA", dirXML)
       cStat = sistNFe.retInutilizacao.infInut.cStat
       NFeMotivo = sistNFe.retInutilizacao.infInut.xMotivo
       NFeDataHora = sistNFe.retInutilizacao.infInut.dhRecbto
       NFeNumeroProtocolo = sistNFe.retInutilizacao.infInut.nProt
       frameAguarde.Visible = False
       DoEvents
       If cStat = 102 Then
          sSQL = "UPDATE TbNFCe SET Inutilizada = 1, Num_OS_VD_Origem = 0 WHERE IdNFProd = " & IdNFProd
          vgDb.Execute sSQL
          MsgBox CStr(cStat) & " - " & NFeMotivo, vbCritical + vbOKOnly, "INUTILIZAÇÃO"
       Else
          MsgBox CStr(cStat) & " - " & NFeMotivo, vbCritical + vbOKOnly, "ERRO - INUTILIZAÇÃO"
       End If
       
       Set sistNFe = Nothing
    End If
End If
Mostrar_NFCe
Grid.SetFocus
End Sub

Private Sub cmdInutilizarTodas_Click()
Dim codPedido As String, nNota As String, CNPJ As String
Dim IdNFProd As Long
Dim varCriterio As String
Dim varIndice As String

If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

    Screen.MousePointer = vbHourglass

    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set r = dbData.OpenRecordset(sSQL)
    NFCeContingencia = r!ContigenciaNFCe
    
    If NFCeContingencia Then
       MsgBox "CONTINGÊNCIA DA NFCE ATIVADA, INUTILIZAÇÃO NÃO PERMITIDA!", vbInformation + vbOKOnly
       GoTo Caifora
    End If
    
    If cboStatus.Text = "" Then Exit Sub
    
    '======================================================= CRITERIOS WHERE
    If cboCriterios.Text = "TODOS" Then
        varCriterio = " WHERE 1=1 "
    ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
        If txtCodPedidoCerto.Text = "" Then
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0 "
        Else
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = " & txtCodPedidoCerto.Text & " "
        End If
    ElseIf cboCriterios.Text = "CLIENTE" Then
        If txtCodCliente.Text = "" Then
            varCriterio = " WHERE TbNFCe.IDCliente = 0"
        Else
            varCriterio = " WHERE TbNFCe.IDCliente = " & txtCodCliente.Text & ""
        End If
    ElseIf cboCriterios.Text = "DATA" Then
        If IsDate(mskData) = True Then
            varCriterio = " WHERE (DataEmissao = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103))"
        Else
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0"
        End If
    ElseIf cboCriterios.Text = "MENSAL" Then
        If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
        
        Dim vIndMes As Integer
        Dim vWhere As String
    
        If cboMes.ListCount = 0 Then
            If cboMes.Text = "Janeiro" Then
                vIndMes = cboMes.ListIndex + 2
            ElseIf cboMes.Text = "Fevereiro" Then
                vIndMes = cboMes.ListIndex + 3
            ElseIf cboMes.Text = "Março" Then
                vIndMes = cboMes.ListIndex + 4
            ElseIf cboMes.Text = "Abril" Then
                vIndMes = cboMes.ListIndex + 5
            ElseIf cboMes.Text = "Maio" Then
                vIndMes = cboMes.ListIndex + 6
            ElseIf cboMes.Text = "Junho" Then
                vIndMes = cboMes.ListIndex + 7
            ElseIf cboMes.Text = "Julho" Then
                vIndMes = cboMes.ListIndex + 8
            ElseIf cboMes.Text = "Agosto" Then
                vIndMes = cboMes.ListIndex + 9
            ElseIf cboMes.Text = "Setembro" Then
                vIndMes = cboMes.ListIndex + 10
            ElseIf cboMes.Text = "Outubro" Then
                vIndMes = cboMes.ListIndex + 11
            ElseIf cboMes.Text = "Novembro" Then
                vIndMes = cboMes.ListIndex + 12
            ElseIf cboMes.Text = "Dezembro" Then
                vIndMes = cboMes.ListIndex + 13
            End If
            
            vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
        Else
            If cboMes.ListIndex = -1 Then
            If cboMes.Text = "Janeiro" Then
                vIndMes = cboMes.ListIndex + 2
            ElseIf cboMes.Text = "Fevereiro" Then
                vIndMes = cboMes.ListIndex + 3
            ElseIf cboMes.Text = "Março" Then
                vIndMes = cboMes.ListIndex + 4
            ElseIf cboMes.Text = "Abril" Then
                vIndMes = cboMes.ListIndex + 5
            ElseIf cboMes.Text = "Maio" Then
                vIndMes = cboMes.ListIndex + 6
            ElseIf cboMes.Text = "Junho" Then
                vIndMes = cboMes.ListIndex + 7
            ElseIf cboMes.Text = "Julho" Then
                vIndMes = cboMes.ListIndex + 8
            ElseIf cboMes.Text = "Agosto" Then
                vIndMes = cboMes.ListIndex + 9
            ElseIf cboMes.Text = "Setembro" Then
                vIndMes = cboMes.ListIndex + 10
            ElseIf cboMes.Text = "Outubro" Then
                vIndMes = cboMes.ListIndex + 11
            ElseIf cboMes.Text = "Novembro" Then
                vIndMes = cboMes.ListIndex + 12
            ElseIf cboMes.Text = "Dezembro" Then
                vIndMes = cboMes.ListIndex + 13
            End If
            
                vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
            Else
                vWhere = "(MONTH(DataEmissao) = " & cboMes.ListIndex + 1 & ") "
            End If
        End If
        
        varCriterio = " WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboAno & ")"
    End If
    
    '====================================================== INDICE
    If cboIndice.Text = "CÓD. PEDIDO" Then
        varIndice = " TbNFCe.Num_OS_VD_Origem "
    ElseIf cboIndice.Text = "CLIENTE" Then
        If cboStatus.Text = "VAZIO" Then
            varIndice = " TbNFCe.Num_OS_VD_Origem "
        Else
            varIndice = " cliente.codigo "
        End If
    ElseIf cboIndice.Text = "EMISSÃO" Then
        varIndice = " TbNFCe.DataEmissao "
    ElseIf cboIndice.Text = "NUM. NFCE" Then
        varIndice = " TbNFCe.IdNFProd "
    Else
        varIndice = " TbNFCe.IdNFProd "
    End If
    
    '==================================================== STATUS
    Dim varStatus As String
    If cboStatus.Text = "TODAS" Then
        varStatus = " "
    ElseIf cboStatus.Text = "ENVIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "NÃO ENVIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "CANCELADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 1 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "INUTILIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 1"
    End If
    
    CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")
    
    sSQL = "SELECT IdNFProd, NumeNota " & _
           "FROM TbNFCe "
    sSQL = sSQL & "" & varCriterio & " " & varStatus & " ORDER BY " & varIndice
        
    Set r = dbData.OpenRecordset(sSQL)
    
    Dim sistNFe As snfe.Util
    Set sistNFe = New snfe.Util
    xCaminhoXML = ""
    
    frameAguarde.Visible = True
    DoEvents
    Do While Not r.EOF
        iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
        nNota = r!NumeNota
        iRetorno = sistNFe.InutilizarNumeracao(Format(Date, "yyyy"), CNPJ, "ERRO AO TRANSMITIR NOTA, PERDA DE SEQUENCIA", nNota, nNota, 1, xCaminhoXML)
        cStat = sistNFe.retInutilizacao.infInut.cStat
        NFeMotivo = sistNFe.retInutilizacao.infInut.xMotivo
        NFeDataHora = sistNFe.retInutilizacao.infInut.dhRecbto
        NFeNumeroProtocolo = sistNFe.retInutilizacao.infInut.nProt
        If cStat = 102 Or cStat = 563 Then
           sSQL = "UPDATE TbNFCe SET Inutilizada = 1, NFCeProtocolo = " & NFeNumeroProtocolo & ", NFCeProtocoloDataHora = '" & NFeDataHora & "', Num_OS_VD_Origem = 0 WHERE IdNFProd = " & r!IdNFProd
           vgDb.Execute sSQL
           'MsgBox CStr(cStat) & " - " & NFeMotivo, vbCritical + vbOKOnly, "INUTILIZAÇÃO"
        Else
           'MsgBox CStr(cStat) & " - " & NFeMotivo, vbCritical + vbOKOnly, "ERRO - INUTILIZAÇÃO"
        End If

        r.MoveNext
    Loop
    
Caifora:
    Set r = Nothing
    frameAguarde.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    Set sistNFe = Nothing
    Mostrar_NFCe
End Sub

Private Sub cmdTransmitir_Click()
ConsultarCPF
ConsultarProdutos
If EncontroErroNFCe = False Then
    Dim sSQL As String, IdNFProd As Long, NomeImpNFCe As String, codPedido As String, retMsg As Integer

    Screen.MousePointer = vbHourglass

    Consultar_Faturas
    Set cCfg = sysConfig("NOME_IMP_NFCE")
    NomeImpNFCe = cCfg.Value
    Set cCfg = Nothing
    
    codPedido = Grid.TextMatrix(Grid.Row, 3)
    
    If NFCeContingencia Then
       MsgBox "CONTINGÊNCIA DA NFCE ATIVADA, ENVIO NÃO PERMITIDO!", vbInformation + vbOKOnly
       GoTo Caifora
    End If
    
    sSQL = "SELECT IdNFProd FROM TbNFCe WHERE IdNFProd  = " & codPedido
    IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)
    If IdNFProd > 0 Then
        frameAguarde.Visible = True
        DoEvents
        iRetorno = TransmitirNFCe(IdNFProd, "1", True, "65")
        If iRetorno Then
           Dim sistNFe As snfe.Util
           Set sistNFe = New snfe.Util
           'Call sistNFe.ImpNFCe(xCaminhoXML, "", "", True, NomeImpNFCe, True, xCaminhoPDF, False, False, 0, "")
        End If
    End If
End If
    
Caifora:
    frameAguarde.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    Mostrar_NFCe
End Sub

Private Sub cmdTransmitirTodas_Click()
Dim varCriterio As String
Dim varIndice As String

If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

    Screen.MousePointer = vbHourglass

    sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
    Set r = dbData.OpenRecordset(sSQL)
    NFCeContingencia = r!ContigenciaNFCe
    
    If NFCeContingencia Then
       MsgBox "CONTINGÊNCIA DA NFCE ATIVADA, ENVIO NÃO PERMITIDO!", vbInformation + vbOKOnly
       GoTo Caifora
    End If
    
    If cboStatus.Text = "" Then Exit Sub
    
    '======================================================= CRITERIOS WHERE
    If cboCriterios.Text = "TODOS" Then
        varCriterio = " WHERE 1=1 "
    ElseIf cboCriterios.Text = "CÓD. PEDIDO" Then
        If txtCodPedidoCerto.Text = "" Then
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0 "
        Else
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = " & txtCodPedidoCerto.Text & " "
        End If
    ElseIf cboCriterios.Text = "CLIENTE" Then
        If txtCodCliente.Text = "" Then
            varCriterio = " WHERE TbNFCe.IDCliente = 0"
        Else
            varCriterio = " WHERE TbNFCe.IDCliente = " & txtCodCliente.Text & ""
        End If
    ElseIf cboCriterios.Text = "DATA" Then
        If IsDate(mskData) = True Then
            varCriterio = " WHERE (DataEmissao = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103))"
        Else
            varCriterio = " WHERE TbNFCe.Num_OS_VD_Origem = 0"
        End If
    ElseIf cboCriterios.Text = "MENSAL" Then
        If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
        
        Dim vIndMes As Integer
        Dim vWhere As String
    
        If cboMes.ListCount = 0 Then
            If cboMes.Text = "Janeiro" Then
                vIndMes = cboMes.ListIndex + 2
            ElseIf cboMes.Text = "Fevereiro" Then
                vIndMes = cboMes.ListIndex + 3
            ElseIf cboMes.Text = "Março" Then
                vIndMes = cboMes.ListIndex + 4
            ElseIf cboMes.Text = "Abril" Then
                vIndMes = cboMes.ListIndex + 5
            ElseIf cboMes.Text = "Maio" Then
                vIndMes = cboMes.ListIndex + 6
            ElseIf cboMes.Text = "Junho" Then
                vIndMes = cboMes.ListIndex + 7
            ElseIf cboMes.Text = "Julho" Then
                vIndMes = cboMes.ListIndex + 8
            ElseIf cboMes.Text = "Agosto" Then
                vIndMes = cboMes.ListIndex + 9
            ElseIf cboMes.Text = "Setembro" Then
                vIndMes = cboMes.ListIndex + 10
            ElseIf cboMes.Text = "Outubro" Then
                vIndMes = cboMes.ListIndex + 11
            ElseIf cboMes.Text = "Novembro" Then
                vIndMes = cboMes.ListIndex + 12
            ElseIf cboMes.Text = "Dezembro" Then
                vIndMes = cboMes.ListIndex + 13
            End If
            
            vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
        Else
            If cboMes.ListIndex = -1 Then
            If cboMes.Text = "Janeiro" Then
                vIndMes = cboMes.ListIndex + 2
            ElseIf cboMes.Text = "Fevereiro" Then
                vIndMes = cboMes.ListIndex + 3
            ElseIf cboMes.Text = "Março" Then
                vIndMes = cboMes.ListIndex + 4
            ElseIf cboMes.Text = "Abril" Then
                vIndMes = cboMes.ListIndex + 5
            ElseIf cboMes.Text = "Maio" Then
                vIndMes = cboMes.ListIndex + 6
            ElseIf cboMes.Text = "Junho" Then
                vIndMes = cboMes.ListIndex + 7
            ElseIf cboMes.Text = "Julho" Then
                vIndMes = cboMes.ListIndex + 8
            ElseIf cboMes.Text = "Agosto" Then
                vIndMes = cboMes.ListIndex + 9
            ElseIf cboMes.Text = "Setembro" Then
                vIndMes = cboMes.ListIndex + 10
            ElseIf cboMes.Text = "Outubro" Then
                vIndMes = cboMes.ListIndex + 11
            ElseIf cboMes.Text = "Novembro" Then
                vIndMes = cboMes.ListIndex + 12
            ElseIf cboMes.Text = "Dezembro" Then
                vIndMes = cboMes.ListIndex + 13
            End If
            
                vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
            Else
                vWhere = "(MONTH(DataEmissao) = " & cboMes.ListIndex + 1 & ") "
            End If
        End If
        
        varCriterio = " WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboAno & ")"
    End If
    
    '====================================================== INDICE
    If cboIndice.Text = "CÓD. PEDIDO" Then
        varIndice = " TbNFCe.Num_OS_VD_Origem "
    ElseIf cboIndice.Text = "CLIENTE" Then
        If cboStatus.Text = "VAZIO" Then
            varIndice = " TbNFCe.Num_OS_VD_Origem "
        Else
            varIndice = " cliente.codigo "
        End If
    ElseIf cboIndice.Text = "EMISSÃO" Then
        varIndice = " TbNFCe.DataEmissao "
    ElseIf cboIndice.Text = "NUM. NFCE" Then
        varIndice = " TbNFCe.IdNFProd "
    Else
        varIndice = " TbNFCe.IdNFProd "
    End If
    
    '==================================================== STATUS
    Dim varStatus As String
    If cboStatus.Text = "TODAS" Then
        varStatus = " "
    ElseIf cboStatus.Text = "ENVIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "NÃO ENVIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "CANCELADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 1 and TbNFCe.NFCeCancelada = 1 and TbNFCe.Inutilizada = 0"
    ElseIf cboStatus.Text = "INUTILIADAS" Then
        varStatus = " and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 1"
    End If
    
    sSQL = "SELECT IdNFProd " & _
           "FROM TbNFCe "
    sSQL = sSQL & "" & varCriterio & " " & varStatus & " ORDER BY " & varIndice
        
    Set r = dbData.OpenRecordset(sSQL)
    
    frameAguarde.Visible = True
    DoEvents
    Do While Not r.EOF
        iRetorno = TransmitirNFCe(r!IdNFProd, "1", True, "65")
        r.MoveNext
    Loop
    
Caifora:
    Set r = Nothing
    frameAguarde.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    Mostrar_NFCe
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
txtCodPedidoCerto.Text = ""
txtCodCliente.Text = ""
txtCodPedido.Text = ""
cboCliente.Text = ""
optDig.Value = True
PreencherCriterios
cboCriterios.ListIndex = 3
PreencherIndice
cboIndice.ListIndex = 0
PreencherStatus
cboStatus.ListIndex = 0
lblData.Visible = True
cmdCal1.Visible = True
mskData.Visible = True
mskData.Text = Format(Date, "dd/mm/yy")
LimprarGridNFCe
cmdConsultar.Visible = False
cmdTransmitir.Visible = False
cmdInutilizar.Visible = False
cmdImprimir.Visible = False
cmdCancelar.Visible = False
cmdDesvincular.Visible = False
cmdExcluir.Visible = False
cmdConsultarTodas.Visible = False
cmdTransmitirTodas.Visible = False
cmdInutilizarTodas.Visible = False
frameAguarde.Visible = False

'nome da caixa
Dim var_Maquina As String
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
StatusBar1.Panels(2).Text = var_Maquina
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

mskData = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub

Private Sub Grid_Click()
Dim i As Long
i = Grid.Row
'mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)

'If lblCodigo.Caption = "" Then Exit Sub
If Grid.TextMatrix(i, 8) = "" Then Exit Sub

If Grid.Rows > 1 Then
'If Grid.TextMatrix(i, 1) <> "000000" Then
    If Grid.TextMatrix(i, 8) = "Inutilizada" Then
        cmdConsultar.Visible = False
        cmdTransmitir.Visible = False
        cmdInutilizar.Visible = False
        cmdImprimir.Enabled = False
        cmdImprimir.Visible = False
        cmdCancelar.Visible = False
        cmdDesvincular.Visible = True
        cmdExcluir.Visible = False
        cmdTransmitirTodas.Visible = False
        cmdInutilizarTodas.Visible = False
    ElseIf Grid.TextMatrix(i, 8) = "Cancelada" Then
        cmdConsultar.Visible = False
        cmdTransmitir.Visible = False
        cmdInutilizar.Visible = False
        cmdImprimir.Enabled = False
        cmdImprimir.Visible = False
        cmdCancelar.Visible = False
        cmdDesvincular.Visible = True
        cmdExcluir.Visible = False
        cmdTransmitirTodas.Visible = False
        cmdInutilizarTodas.Visible = False
        cmdChave2.Visible = True
        cmdEnviarPDF.Visible = True
        cmdEnviarArquivo.Visible = True
    ElseIf Grid.TextMatrix(i, 8) = "Enviada" Then
        cmdConsultar.Visible = True
        cmdTransmitir.Visible = False
        cmdInutilizar.Visible = False
        cmdImprimir.Enabled = True
        cmdImprimir.Visible = True
        cmdCancelar.Visible = True
        cmdDesvincular.Visible = True
        cmdExcluir.Visible = False
        cmdTransmitirTodas.Visible = False
        cmdInutilizarTodas.Visible = False
        cmdChave2.Visible = True
        cmdEnviarPDF.Visible = True
        cmdEnviarArquivo.Visible = True
    ElseIf Grid.TextMatrix(i, 8) = "Contingência" Then
        cmdConsultar.Visible = True
        cmdTransmitir.Visible = True
        cmdInutilizar.Visible = True
        cmdImprimir.Enabled = True
        cmdCancelar.Visible = False
        cmdDesvincular.Visible = True
        cmdExcluir.Visible = True
        If cboStatus.Text = "NÃO ENVIADAS" Then cmdTransmitirTodas.Visible = True
        If cboStatus.Text = "NÃO ENVIADAS" Then cmdInutilizarTodas.Visible = True
    Else
        cmdConsultar.Visible = True
        cmdTransmitir.Visible = True
        cmdInutilizar.Visible = True
        cmdImprimir.Enabled = False
        cmdImprimir.Visible = True
        cmdCancelar.Visible = False
        cmdDesvincular.Visible = True
        cmdExcluir.Visible = True
        'cmdTransmitirTodas.Visible = True
        'cmdInutilizarTodas.Visible = True
    End If
'Else
'        cmdConsultar.visible = False
'        cmdTransmitir.visible = False
'        cmdInutilizar.visible = False
'        cmdImprimir.visible = False
'        cmdCancelar.visible = False
'        cmdDesvincular.visible = False
'        cmdExcluir.visible = False
End If
'


End Sub

Private Sub mskData_GotFocus()
SelectControl mskData
End Sub


Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
End Sub


Private Sub mskData_LostFocus()
If Not IsDate(mskData.Text) Then
   mskData.Mask = ""
   mskData.Text = ""
End If
End Sub
Private Sub cboCliente_Click()
   On Error GoTo TrataErro
   
   If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub CboCliente_GotFocus()
cboCliente.Clear

sSQL = "SELECT * FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCliente.AddItem r("nome")
   cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CboCliente_LostFocus()
cboCliente_Click
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer
Dim vTexto As String
vTexto = cboAno.Text

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next
cboAno.Text = vTexto
moCombo.AttachTo cboAno
End Sub
Private Sub cboMes_GotFocus()
Dim vTexto As String
vTexto = cboMes.Text
cboMes.Clear

cboMes.AddItem "Janeiro"
cboMes.AddItem "Fevereiro"
cboMes.AddItem "Março"
cboMes.AddItem "Abril"
cboMes.AddItem "Maio"
cboMes.AddItem "Junho"
cboMes.AddItem "Julho"
cboMes.AddItem "Agosto"
cboMes.AddItem "Setembro"
cboMes.AddItem "Outubro"
cboMes.AddItem "Novembro"
cboMes.AddItem "Dezembro"
cboMes.Text = vTexto
moCombo.AttachTo cboMes
End Sub


Private Sub cboMes_LostFocus()
cboAno.SetFocus
End Sub

Private Sub optDig_Click()
txtCodPedido.Clear
End Sub

Private Sub optEsc_Click()
txtCodPedido.Clear
End Sub


Private Sub txtCodPedido_Change()
If optDig.Value = True Then
    If txtCodPedido.Text <> "" Then
        txtCodPedidoCerto.Text = txtCodPedido.Text
    Else
        txtCodPedidoCerto.Text = ""
    End If
Else
    If txtCodPedido.Text <> "" Then
    txtCodPedidoCerto.Text = Mid(txtCodPedido.Text, 1, InStr(1, txtCodPedido.Text, "->", vbTextCompare) - 1)
    End If
End If
End Sub

Private Sub txtCodPedido_Click()
If optDig.Value = True Then
    If txtCodPedido.Text <> "" Then
    txtCodPedidoCerto.Text = txtCodPedido.Text
    End If
Else
    If txtCodPedido.Text <> "" Then
    txtCodPedidoCerto.Text = Mid(txtCodPedido.Text, 1, InStr(1, txtCodPedido.Text, "->", vbTextCompare) - 1)
    End If
End If
txtCodPedido_LostFocus
End Sub

Private Sub txtCodPedido_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   txtCodPedido.Clear
   
   sSQL = "SELECT top 50 * FROM pedidos ORDER BY cod_pedido DESC;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF Then
      txtCodPedido.AddItem "NENHUM PEDIDO"
   Else
      Do While Not r.EOF

        If optDig.Value = True Then
            txtCodPedido.Clear
        Else
            txtCodPedido.AddItem Format(r("cod_pedido"), "000000") & " -> " & Format(r("data_compra"), "dd/mm/yy") & " -> " & Format(r("total"), ocMONEY)
        End If
         
         r.MoveNext
      Loop
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

If optEsc.Value = True Then
   txtCodPedido.ListIndex = 0
End If
End Sub


Private Sub txtCodPedido_LostFocus()
If txtCodPedido.Text = "" Then Exit Sub
'Mostrar_Pedido
End Sub

Private Sub txtCodPedido_Validate(Cancel As Boolean)
If txtCodPedido.Text = "" Then
    If optEsc.Value = True Then
    txtCodPedidoCerto.Text = Mid(txtCodPedido.Text, 1, InStr(1, txtCodPedido.Text, "->", vbTextCompare) - 1)
    End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
HabilitaObjetosVenda False
Set moCombo = Nothing
End Sub


