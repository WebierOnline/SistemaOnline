VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form NFe_Completa 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NFe - Nota Fiscal Eletronica"
   ClientHeight    =   9885
   ClientLeft      =   735
   ClientTop       =   1455
   ClientWidth     =   18210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Frm_NF"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9885
   ScaleWidth      =   18210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Frm_NF 
      Height          =   8655
      Left            =   60
      TabIndex        =   70
      Top             =   840
      Width           =   18075
      _ExtentX        =   31882
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   5292
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "NFe_Completa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSalvar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNovo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Tab_Produtos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Tab_Totais"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmDestinatario"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmNota"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancelar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "NOTAS FISCAIS"
      TabPicture(1)   =   "NFe_Completa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEnviarPDF"
      Tab(1).Control(1)=   "cmdEnviarXML"
      Tab(1).Control(2)=   "cmdEspelho"
      Tab(1).Control(3)=   "cmdEditar"
      Tab(1).Control(4)=   "cmdCartaCorrecao"
      Tab(1).Control(5)=   "cmdInutilizar"
      Tab(1).Control(6)=   "cmdDuplicar"
      Tab(1).Control(7)=   "cmdConsultar"
      Tab(1).Control(8)=   "cmdCancelarNota"
      Tab(1).Control(9)=   "cmdTransmitir"
      Tab(1).Control(10)=   "cmdImprimir"
      Tab(1).Control(11)=   "cmdCopiarChave"
      Tab(1).Control(12)=   "GridNotas"
      Tab(1).Control(13)=   "Frame4"
      Tab(1).Control(14)=   "Frame2"
      Tab(1).Control(15)=   "picAguarde"
      Tab(1).Control(16)=   "frmCorreção"
      Tab(1).Control(17)=   "txtCodObservacao"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "PEDIDOS"
      TabPicture(2)   =   "NFe_Completa.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblQuantPedidos"
      Tab(2).Control(1)=   "cmdConverterNFe"
      Tab(2).Control(2)=   "GridPedidos"
      Tab(2).Control(3)=   "frmFiltrosPedidos"
      Tab(2).Control(4)=   "picAguarde2"
      Tab(2).ControlCount=   5
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   14700
         TabIndex        =   48
         Top             =   1740
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "NFe_Completa.frx":0054
         PICN            =   "NFe_Completa.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox picAguarde2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -68220
         Picture         =   "NFe_Completa.frx":1E02
         ScaleHeight     =   1095
         ScaleWidth      =   2895
         TabIndex        =   233
         Top             =   3300
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtCodObservacao 
         Height          =   315
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   6900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame frmCorreção 
         BackColor       =   &H0080C0FF&
         Caption         =   "Carta de Correção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -73620
         TabIndex        =   216
         Top             =   1380
         Visible         =   0   'False
         Width           =   12615
         Begin VB.TextBox txtCorrecao 
            Height          =   375
            Left            =   180
            TabIndex        =   217
            Top             =   540
            Width           =   12255
         End
         Begin ChamaleonBtn.chameleonButton cmdCCeImprimir 
            Height          =   375
            Left            =   3600
            TabIndex        =   218
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "NFe_Completa.frx":2E3A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCCeSalvar 
            Height          =   375
            Left            =   180
            TabIndex        =   219
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Salvar"
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
            MICON           =   "NFe_Completa.frx":2E56
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCCeExcluir 
            Height          =   375
            Left            =   2460
            TabIndex        =   220
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Excluir"
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
            MICON           =   "NFe_Completa.frx":2E72
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCCeTransmitir 
            Height          =   375
            Left            =   1320
            TabIndex        =   221
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Transmitir"
            ENAB            =   0   'False
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
            MICON           =   "NFe_Completa.frx":2E8E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Correcao 
            Height          =   2415
            Left            =   180
            TabIndex        =   222
            Top             =   1380
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   4260
            _Version        =   393216
            TextStyleFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdFecharCCe 
            Height          =   375
            Left            =   11460
            TabIndex        =   224
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Fechar"
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
            MICON           =   "NFe_Completa.frx":2EAA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Correção"
            Height          =   195
            Left            =   180
            TabIndex        =   223
            Top             =   300
            Width           =   645
         End
      End
      Begin VB.PictureBox picAguarde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -68100
         Picture         =   "NFe_Completa.frx":2EC6
         ScaleHeight     =   1095
         ScaleWidth      =   2895
         TabIndex        =   215
         Top             =   3000
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "Totais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1155
         Left            =   -61920
         TabIndex        =   201
         Top             =   6360
         Width           =   3435
         Begin VB.Label lblQuantInutilizada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1410
            TabIndex        =   213
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblTotalInutilizada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2010
            TabIndex        =   212
            Top             =   900
            Width           =   1245
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inutilizadas:"
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
            Left            =   330
            TabIndex        =   211
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lblQuantEnviada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1410
            TabIndex        =   210
            Top             =   180
            Width           =   525
         End
         Begin VB.Label lblTotalEnviada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2010
            TabIndex        =   209
            Top             =   180
            Width           =   1245
         End
         Begin VB.Label lblTotalNaoEnviada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2010
            TabIndex        =   208
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label lblQuantNaoEnviada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1410
            TabIndex        =   207
            Top             =   420
            Width           =   525
         End
         Begin VB.Label lblQuantCancelada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1410
            TabIndex        =   206
            Top             =   660
            Width           =   525
         End
         Begin VB.Label lblTotalCancelada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2010
            TabIndex        =   205
            Top             =   660
            Width           =   1245
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Enviadas:"
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
            TabIndex        =   204
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Não Enviadas:"
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
            TabIndex        =   203
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Canceladas:"
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
            Left            =   330
            TabIndex        =   202
            Top             =   660
            Width           =   1065
         End
      End
      Begin VB.Frame frmFiltrosPedidos 
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
         Height          =   1155
         Left            =   -74880
         TabIndex        =   159
         Top             =   7440
         Width           =   16395
         Begin VB.Frame Frame9 
            Caption         =   "Filtrar por:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   180
            TabIndex        =   176
            Top             =   240
            Width           =   4035
            Begin VB.ComboBox cboIndicePedidos 
               Height          =   315
               Left            =   960
               TabIndex        =   177
               Top             =   300
               Width           =   2715
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Escolher:"
               Height          =   195
               Left            =   180
               TabIndex        =   178
               Top             =   360
               Width           =   660
            End
         End
         Begin VB.Frame Frame8 
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
            Height          =   735
            Left            =   4320
            TabIndex        =   160
            Top             =   240
            Width           =   5535
            Begin VB.ComboBox cboAnoPedidos 
               Height          =   315
               Left            =   2340
               Sorted          =   -1  'True
               TabIndex        =   165
               Top             =   240
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMesPedidos 
               Height          =   315
               Left            =   540
               TabIndex        =   164
               Top             =   240
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboClientePedidos 
               Height          =   315
               Left            =   720
               TabIndex        =   163
               Top             =   240
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.TextBox txtCodClientePedidos 
               Height          =   315
               Left            =   4680
               TabIndex        =   162
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtConCodPedido 
               Height          =   315
               Left            =   1020
               TabIndex        =   161
               Top             =   240
               Visible         =   0   'False
               Width           =   1875
            End
            Begin MSMask.MaskEdBox mskFinalPedidos 
               Height          =   315
               Left            =   2640
               TabIndex        =   166
               Top             =   260
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCalPedidos2 
               Height          =   315
               Left            =   3660
               TabIndex        =   167
               Top             =   260
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
               MICON           =   "NFe_Completa.frx":3EFE
               PICN            =   "NFe_Completa.frx":3F1A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskInicialPedidos 
               Height          =   315
               Left            =   720
               TabIndex        =   168
               Top             =   260
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCalPedidos1 
               Height          =   315
               Left            =   1740
               TabIndex        =   169
               Top             =   260
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
               MICON           =   "NFe_Completa.frx":62FC
               PICN            =   "NFe_Completa.frx":6318
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblAnoPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1980
               TabIndex        =   175
               Top             =   240
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblMesPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   174
               Top             =   240
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblClientePedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   120
               TabIndex        =   173
               Top             =   240
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblInicialPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   172
               Top             =   260
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lblFinalPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Final:"
               Height          =   195
               Left            =   2220
               TabIndex        =   171
               Top             =   260
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblConsCodPedido 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. Pedido:"
               Height          =   195
               Left            =   60
               TabIndex        =   170
               Top             =   240
               Visible         =   0   'False
               Width           =   915
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
            Height          =   495
            Left            =   9960
            TabIndex        =   179
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
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
            MICON           =   "NFe_Completa.frx":86FA
            PICN            =   "NFe_Completa.frx":8716
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
      Begin VB.Frame Frame4 
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
         Left            =   -74940
         TabIndex        =   145
         Top             =   7440
         Width           =   16455
         Begin VB.Frame Frame7 
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
            Height          =   675
            Left            =   4140
            TabIndex        =   149
            Top             =   240
            Width           =   5535
            Begin VB.ComboBox cboConNotaMes 
               Height          =   315
               Left            =   960
               TabIndex        =   150
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox cboConNotaAno 
               Height          =   315
               Left            =   2460
               TabIndex        =   151
               Top             =   240
               Visible         =   0   'False
               Width           =   1215
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal1 
               Height          =   315
               Left            =   1920
               TabIndex        =   185
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "NFe_Completa.frx":8FF0
               PICN            =   "NFe_Completa.frx":900C
               PICH            =   "NFe_Completa.frx":B35F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal2 
               Height          =   315
               Left            =   3300
               TabIndex        =   186
               Tag             =   "Calendario"
               Top             =   240
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
               MICON           =   "NFe_Completa.frx":D6B2
               PICN            =   "NFe_Completa.frx":D6CE
               PICH            =   "NFe_Completa.frx":FA21
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txtConNotaCodCliente 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5040
               TabIndex        =   153
               Top             =   60
               Width           =   495
            End
            Begin VB.ComboBox cboConNotaCliente 
               Height          =   315
               Left            =   1080
               TabIndex        =   152
               Top             =   240
               Visible         =   0   'False
               Width           =   4305
            End
            Begin MSMask.MaskEdBox mskConNotaFinal 
               Height          =   315
               Left            =   2340
               TabIndex        =   154
               Top             =   240
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskConNotaInicial 
               Height          =   315
               Left            =   960
               TabIndex        =   155
               Top             =   240
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin VB.Label lblConNotaAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   180
               TabIndex        =   156
               Top             =   300
               Visible         =   0   'False
               Width           =   330
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Filtrar por:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   60
            TabIndex        =   146
            Top             =   240
            Width           =   4035
            Begin VB.ComboBox cboFiltroNota 
               Height          =   315
               Left            =   960
               TabIndex        =   147
               Top             =   240
               Width           =   2715
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Escolher:"
               Height          =   195
               Left            =   180
               TabIndex        =   148
               Top             =   300
               Width           =   660
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConNotas 
            Height          =   495
            Left            =   9780
            TabIndex        =   157
            Top             =   420
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   873
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
            MICON           =   "NFe_Completa.frx":11D74
            PICN            =   "NFe_Completa.frx":11D90
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdImprimirConsulta 
            Height          =   495
            Left            =   11340
            TabIndex        =   188
            Top             =   420
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   873
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
            MICON           =   "NFe_Completa.frx":1266A
            PICN            =   "NFe_Completa.frx":12686
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
      Begin VB.Frame frmNota 
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   120
         TabIndex        =   144
         Top             =   1260
         Width           =   14535
         Begin VB.TextBox txtSerie 
            Height          =   315
            Left            =   840
            MaxLength       =   50
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   540
            Width           =   680
         End
         Begin ChamaleonBtn.chameleonButton cmdDuplicarCFOP 
            Height          =   315
            Left            =   7380
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Coloca o mesmo CFOP em todos os itens da nota fiscal"
            Top             =   540
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "..."
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
            MICON           =   "NFe_Completa.frx":14418
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboFinalidade 
            Height          =   315
            ItemData        =   "NFe_Completa.frx":14434
            Left            =   2940
            List            =   "NFe_Completa.frx":14436
            TabIndex        =   9
            Top             =   540
            Width           =   1755
         End
         Begin VB.ComboBox cboTipoNota 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Top             =   540
            Width           =   1335
         End
         Begin VB.TextBox txtNatureza 
            Height          =   315
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   540
            Width           =   3300
         End
         Begin VB.ComboBox cboNatureza 
            Height          =   315
            Left            =   6660
            TabIndex        =   11
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtNumNota 
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   540
            Width           =   645
         End
         Begin VB.ComboBox cboDestOperacao 
            Height          =   315
            Left            =   4740
            TabIndex        =   10
            Top             =   540
            Width           =   1875
         End
         Begin ChamaleonBtn.chameleonButton cmdCal2 
            Height          =   315
            Left            =   13320
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   540
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
            MICON           =   "NFe_Completa.frx":14438
            PICN            =   "NFe_Completa.frx":14454
            PICH            =   "NFe_Completa.frx":167A7
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
            Left            =   12000
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   540
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
            MICON           =   "NFe_Completa.frx":18AFA
            PICN            =   "NFe_Completa.frx":18B16
            PICH            =   "NFe_Completa.frx":1AE69
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
            Left            =   13680
            TabIndex        =   18
            Top             =   540
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSaida 
            Height          =   315
            Left            =   12360
            TabIndex        =   16
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEmissao 
            Height          =   315
            Left            =   11040
            TabIndex        =   14
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
            Height          =   195
            Index           =   12
            Left            =   13680
            TabIndex        =   253
            Top             =   300
            Width           =   345
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Saida"
            Height          =   195
            Index           =   11
            Left            =   12360
            TabIndex        =   252
            Top             =   300
            Width           =   660
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Emissão"
            Height          =   195
            Index           =   10
            Left            =   11040
            TabIndex        =   251
            Top             =   300
            Width           =   840
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Natureza da Operação"
            Height          =   195
            Index           =   9
            Left            =   6660
            TabIndex        =   250
            Top             =   300
            Width           =   1620
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destino"
            Height          =   195
            Index           =   8
            Left            =   4740
            TabIndex        =   249
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finalidade da Emissão"
            Height          =   195
            Index           =   7
            Left            =   2940
            TabIndex        =   248
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Nota"
            Height          =   195
            Index           =   6
            Left            =   1560
            TabIndex        =   247
            Top             =   300
            Width           =   930
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Série"
            Height          =   195
            Index           =   5
            Left            =   840
            TabIndex        =   246
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NF Num."
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   245
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.Frame frmDestinatario 
         Caption         =   "Destinatário"
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
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   14535
         Begin VB.TextBox txtAliqUFDest 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   14040
            MaxLength       =   10
            TabIndex        =   187
            ToolTipText     =   "Aliquota Dest"
            Top             =   480
            Width           =   390
         End
         Begin VB.ComboBox cboTipoDest 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox cboConsumidorFinal 
            Height          =   315
            Left            =   10680
            TabIndex        =   4
            Top             =   480
            Width           =   1875
         End
         Begin VB.ComboBox cboTipoContribuinte 
            Height          =   315
            Left            =   8100
            TabIndex        =   3
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox TxtCodCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   7020
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.ComboBox CboCliente 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2220
            TabIndex        =   2
            Top             =   480
            Width           =   5835
         End
         Begin ChamaleonBtn.chameleonButton cmdConsultarCliente 
            Height          =   315
            Left            =   12600
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   480
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MICON           =   "NFe_Completa.frx":1D1BC
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
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consumidor Final"
            Height          =   195
            Index           =   3
            Left            =   10680
            TabIndex        =   244
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Contribuinte"
            Height          =   195
            Index           =   2
            Left            =   8100
            TabIndex        =   243
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente / Destinatário"
            Height          =   195
            Index           =   1
            Left            =   2220
            TabIndex        =   242
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Destinatário"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1425
         End
      End
      Begin TabDlg.SSTab Tab_Totais 
         Height          =   1095
         Left            =   120
         TabIndex        =   74
         Top             =   7380
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   1931
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   3528
         TabCaption(0)   =   "Totais da Nota"
         TabPicture(0)   =   "NFe_Completa.frx":1D1D8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label21"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label20"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label19"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label16"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label14"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label27"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label37"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label38"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label41"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label44"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label45"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtValorFrete"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtBaseICMS"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtBaseICMSST"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtValorSeguro"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtValorOutrasDespesas"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtValorICMS"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtValorICMSST"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtValorIPI"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtValorDesconto"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtTotaldaNota"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtTotaldosProdutos"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Outros Tributos"
         TabPicture(1)   =   "NFe_Completa.frx":1D1F4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Retenção de Tributos"
         TabPicture(2)   =   "NFe_Completa.frx":1D210
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "Interestadual"
         TabPicture(3)   =   "NFe_Completa.frx":1D22C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin VB.TextBox txtTotaldosProdutos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   315
            Left            =   5160
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox txtTotaldaNota 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   12960
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox txtValorDesconto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9120
            MaxLength       =   50
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtValorIPI 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11700
            MaxLength       =   50
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtValorICMSST 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3900
            MaxLength       =   50
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtValorICMS 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1380
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtValorOutrasDespesas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10380
            MaxLength       =   50
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   600
            Width           =   1305
         End
         Begin VB.TextBox txtValorSeguro 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7860
            MaxLength       =   50
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtBaseICMSST 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtBaseICMS 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   105
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtValorFrete 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Produtos"
            Height          =   195
            Left            =   5160
            TabIndex        =   85
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total da Nota"
            Height          =   195
            Left            =   12960
            TabIndex        =   84
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto"
            Height          =   195
            Left            =   9120
            TabIndex        =   83
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do IPI"
            Height          =   195
            Left            =   11700
            TabIndex        =   82
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total ICMS ST"
            Height          =   195
            Left            =   3900
            TabIndex        =   81
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS"
            Height          =   195
            Left            =   1380
            TabIndex        =   80
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Frete"
            Height          =   195
            Left            =   6600
            TabIndex        =   79
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outras Despesas"
            Height          =   195
            Left            =   10380
            TabIndex        =   78
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seguro"
            Height          =   195
            Left            =   7860
            TabIndex        =   77
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base ICMS ST"
            Height          =   195
            Left            =   2640
            TabIndex        =   76
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base ICMS"
            Height          =   195
            Left            =   105
            TabIndex        =   75
            Top             =   360
            Width           =   795
         End
      End
      Begin TabDlg.SSTab Tab_Produtos 
         Height          =   4995
         Left            =   120
         TabIndex        =   86
         Top             =   2340
         Width           =   17865
         _ExtentX        =   31512
         _ExtentY        =   8811
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   467
         TabMaxWidth     =   3351
         TabCaption(0)   =   "Produtos"
         TabPicture(0)   =   "NFe_Completa.frx":1D248
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmItens"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Transporte"
         TabPicture(1)   =   "NFe_Completa.frx":1D264
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label71"
         Tab(1).Control(1)=   "Tab_transp"
         Tab(1).Control(2)=   "cboModFrete"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Cobrança"
         TabPicture(2)   =   "NFe_Completa.frx":1D280
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label15"
         Tab(2).Control(1)=   "Label67"
         Tab(2).Control(2)=   "cboIndicadorPagamento"
         Tab(2).Control(3)=   "frmFatura"
         Tab(2).Control(4)=   "frmDuplicata"
         Tab(2).Control(5)=   "cboFormaPgto"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Informações"
         TabPicture(3)   =   "NFe_Completa.frx":1D29C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Tab_Info"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "DANFe"
         TabPicture(4)   =   "NFe_Completa.frx":1D2B8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cboFormatoDANFe"
         Tab(4).Control(1)=   "cboTipoEmissao"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Exportação e Compra"
         TabPicture(5)   =   "NFe_Completa.frx":1D2D4
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         TabCaption(6)   =   "Devolução"
         TabPicture(6)   =   "NFe_Completa.frx":1D2F0
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "frmDevolucao"
         Tab(6).ControlCount=   1
         Begin VB.Frame frmDevolucao 
            Caption         =   "Devolução"
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
            Left            =   -74880
            TabIndex        =   189
            Top             =   360
            Width           =   14295
            Begin VB.TextBox txtChaveReferenciada 
               Height          =   315
               Left            =   120
               TabIndex        =   190
               Top             =   480
               Width           =   13695
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Chave de Acesso - Nota de entrada"
               Height          =   195
               Left            =   120
               TabIndex        =   191
               Top             =   240
               Width           =   2550
            End
         End
         Begin VB.ComboBox cboFormaPgto 
            Height          =   315
            Left            =   -71640
            TabIndex        =   183
            Top             =   660
            Width           =   3135
         End
         Begin VB.Frame frmDuplicata 
            Caption         =   "Duplicata"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   -74880
            TabIndex        =   101
            Top             =   1980
            Visible         =   0   'False
            Width           =   14235
            Begin VB.TextBox txtIntervaloDup 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4500
               MaxLength       =   50
               TabIndex        =   58
               Top             =   480
               Width           =   720
            End
            Begin VB.TextBox txtValorParcDup 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   6555
               MaxLength       =   50
               TabIndex        =   60
               Top             =   480
               Width           =   1320
            End
            Begin VB.TextBox txtTotalDup 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1755
               MaxLength       =   50
               TabIndex        =   56
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtNumParcDup 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3375
               MaxLength       =   50
               TabIndex        =   57
               Top             =   480
               Width           =   1080
            End
            Begin VB.TextBox txtNumDup 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   55
               Top             =   480
               Width           =   1560
            End
            Begin ChamaleonBtn.chameleonButton cmdCalDuplic 
               Height          =   315
               Left            =   6240
               TabIndex        =   102
               Tag             =   "Calendario"
               Top             =   480
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
               MICON           =   "NFe_Completa.frx":1D30C
               PICN            =   "NFe_Completa.frx":1D328
               PICH            =   "NFe_Completa.frx":1F67B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskInicioDup 
               Height          =   315
               Left            =   5280
               TabIndex        =   59
               Top             =   480
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCriarDuplicata 
               Height          =   315
               Left            =   7980
               TabIndex        =   234
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Criar Duplicatas"
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
               MICON           =   "NFe_Completa.frx":219CE
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverDuplicatas 
               Height          =   315
               Left            =   9360
               TabIndex        =   235
               Top             =   480
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Excluir Duplicatas"
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
               MICON           =   "NFe_Completa.frx":219EA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Duplicata 
               Height          =   1875
               Left            =   120
               TabIndex        =   236
               Top             =   900
               Width           =   10695
               _ExtentX        =   18865
               _ExtentY        =   3307
               _Version        =   393216
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio:"
               Height          =   195
               Left            =   5280
               TabIndex        =   108
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Intervalo"
               Height          =   195
               Left            =   4500
               TabIndex        =   107
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor da Parcela"
               Height          =   195
               Left            =   6555
               TabIndex        =   106
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   195
               Left            =   1755
               TabIndex        =   105
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant. Parc."
               Height          =   195
               Left            =   3375
               TabIndex        =   104
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número/Doc."
               Height          =   195
               Left            =   120
               TabIndex        =   103
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame frmFatura 
            Caption         =   "Fatura"
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
            Left            =   -74880
            TabIndex        =   96
            Top             =   1020
            Width           =   14235
            Begin VB.TextBox txtNumFatura 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   51
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtDescFatura 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3375
               MaxLength       =   50
               TabIndex        =   53
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtSubtotalFatura 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1755
               MaxLength       =   50
               TabIndex        =   52
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtTotalFatura 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   4995
               MaxLength       =   50
               TabIndex        =   54
               Top             =   480
               Width           =   1560
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número"
               Height          =   195
               Left            =   120
               TabIndex        =   100
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto"
               Height          =   195
               Left            =   3375
               TabIndex        =   99
               Top             =   240
               Width           =   690
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal"
               Height          =   195
               Left            =   1755
               TabIndex        =   98
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   195
               Left            =   4995
               TabIndex        =   97
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.ComboBox cboTipoEmissao 
            Height          =   315
            Left            =   -73020
            TabIndex        =   64
            Top             =   720
            Width           =   2595
         End
         Begin VB.ComboBox cboFormatoDANFe 
            Height          =   315
            Left            =   -74880
            TabIndex        =   63
            Top             =   720
            Width           =   1875
         End
         Begin VB.ComboBox cboIndicadorPagamento 
            Height          =   315
            Left            =   -74820
            TabIndex        =   50
            Top             =   660
            Width           =   3135
         End
         Begin VB.ComboBox cboModFrete 
            Height          =   315
            Left            =   -74880
            TabIndex        =   49
            Top             =   600
            Width           =   5055
         End
         Begin VB.Frame frmItens 
            Caption         =   "Itens da Nota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   120
            TabIndex        =   87
            Top             =   300
            Width           =   17655
            Begin VB.CheckBox chkpRedBC 
               Caption         =   "RedBC"
               Height          =   195
               Left            =   10920
               TabIndex        =   256
               Top             =   4260
               Width           =   1035
            End
            Begin VB.CheckBox chkICMSST 
               Caption         =   "ICMSST"
               Height          =   195
               Left            =   9780
               TabIndex        =   255
               Top             =   4260
               Width           =   1035
            End
            Begin VB.CheckBox chkIPI 
               Caption         =   "IPI"
               Height          =   195
               Left            =   9060
               TabIndex        =   254
               Top             =   4260
               Width           =   615
            End
            Begin VB.TextBox txtFrete 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9120
               MaxLength       =   10
               TabIndex        =   23
               Top             =   480
               Width           =   825
            End
            Begin VB.TextBox txtSeguro 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9960
               MaxLength       =   10
               TabIndex        =   24
               Top             =   480
               Width           =   825
            End
            Begin VB.TextBox txtOutrosItem 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10800
               MaxLength       =   10
               TabIndex        =   25
               Top             =   480
               Width           =   825
            End
            Begin VB.ComboBox cboDescricao 
               Height          =   315
               Left            =   1740
               TabIndex        =   20
               Top             =   480
               Width           =   5535
            End
            Begin VB.TextBox txtCodBarra 
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtCodProduto 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6420
               TabIndex        =   89
               Top             =   240
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtQuant 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8280
               MaxLength       =   10
               TabIndex        =   22
               Top             =   480
               Width           =   810
            End
            Begin VB.TextBox txtSubTotal 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
               Height          =   315
               Left            =   12660
               Locked          =   -1  'True
               MaxLength       =   8
               TabIndex        =   27
               Top             =   480
               Width           =   1080
            End
            Begin VB.TextBox txtValor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7260
               MaxLength       =   8
               TabIndex        =   21
               Top             =   480
               Width           =   960
            End
            Begin VB.TextBox txtDesc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   11640
               MaxLength       =   10
               TabIndex        =   26
               Top             =   480
               Width           =   990
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   5040
               TabIndex        =   88
               Top             =   1800
               Visible         =   0   'False
               Width           =   810
            End
            Begin MSFlexGridLib.MSFlexGrid GridNotasItens 
               Height          =   3315
               Left            =   120
               TabIndex        =   29
               Top             =   840
               Width           =   17415
               _ExtentX        =   30718
               _ExtentY        =   5847
               _Version        =   393216
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarItem 
               Height          =   315
               Left            =   13800
               TabIndex        =   28
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "A&dicionar"
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
               MICON           =   "NFe_Completa.frx":21A06
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverItem 
               Height          =   315
               Left            =   15000
               TabIndex        =   30
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "R&emover"
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
               MICON           =   "NFe_Completa.frx":21A22
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConsultarNCM 
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   4200
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Consultar NCM pela Descrição"
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
               MICON           =   "NFe_Completa.frx":21A3E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConsultarProduto 
               Height          =   255
               Left            =   6900
               TabIndex        =   34
               Top             =   4200
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Atualizar Produto"
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
               MICON           =   "NFe_Completa.frx":21A5A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConsultaNCMean 
               Height          =   255
               Left            =   2700
               TabIndex        =   32
               Top             =   4200
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Consultar NCM pelo EAN"
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
               MICON           =   "NFe_Completa.frx":21A76
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRecalcular 
               Height          =   255
               Left            =   14460
               TabIndex        =   35
               Top             =   4200
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Recalcular"
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
               MICON           =   "NFe_Completa.frx":21A92
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdConsultarCest 
               Height          =   255
               Left            =   5280
               TabIndex        =   33
               Top             =   4200
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Consultar Cest"
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
               MICON           =   "NFe_Completa.frx":21AAE
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label txtOutros 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Outros"
               Height          =   195
               Left            =   10800
               TabIndex        =   241
               Top             =   240
               Width           =   465
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Seguro"
               Height          =   195
               Left            =   9960
               TabIndex        =   240
               Top             =   240
               Width           =   510
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frete"
               Height          =   195
               Left            =   9120
               TabIndex        =   239
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               Height          =   195
               Left            =   1740
               TabIndex        =   95
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblCodFabrica 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. de Barra"
               Height          =   195
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal"
               Height          =   195
               Left            =   12660
               TabIndex        =   93
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   7260
               TabIndex        =   92
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde"
               Height          =   195
               Left            =   8280
               TabIndex        =   91
               Top             =   240
               Width           =   345
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto"
               Height          =   195
               Left            =   11640
               TabIndex        =   90
               Top             =   240
               Width           =   690
            End
         End
         Begin TabDlg.SSTab Tab_Info 
            Height          =   4515
            Left            =   -74880
            TabIndex        =   109
            Top             =   360
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   7964
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabMaxWidth     =   5292
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Informações Complementares"
            TabPicture(0)   =   "NFe_Completa.frx":21ACA
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "cmdRemoverOBS"
            Tab(0).Control(1)=   "cmdAdicionarOBS"
            Tab(0).Control(2)=   "txtInfComple"
            Tab(0).Control(3)=   "cboObservacao"
            Tab(0).Control(4)=   "txtCodOBS"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).ControlCount=   5
            TabCaption(1)   =   "Informações Adicionais"
            TabPicture(1)   =   "NFe_Completa.frx":21AE6
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "txtInfAdicionais"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.TextBox txtCodOBS 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   -65040
               Locked          =   -1  'True
               TabIndex        =   231
               TabStop         =   0   'False
               Top             =   300
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.ComboBox cboObservacao 
               Height          =   315
               Left            =   -74880
               TabIndex        =   228
               Top             =   480
               Width           =   10515
            End
            Begin VB.TextBox txtInfAdicionais 
               Height          =   4005
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   62
               Top             =   420
               Width           =   13920
            End
            Begin VB.TextBox txtInfComple 
               Height          =   3585
               Left            =   -74880
               MultiLine       =   -1  'True
               TabIndex        =   61
               Top             =   840
               Width           =   13920
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarOBS 
               Height          =   315
               Left            =   -64320
               TabIndex        =   229
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Adicionar"
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
               MICON           =   "NFe_Completa.frx":21B02
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverOBS 
               Height          =   315
               Left            =   -63120
               TabIndex        =   230
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Remover"
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
               MICON           =   "NFe_Completa.frx":21B1E
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
         Begin TabDlg.SSTab Tab_transp 
            Height          =   2295
            Left            =   -74940
            TabIndex        =   110
            Top             =   1140
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   4048
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabMaxWidth     =   2646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Transportadora"
            TabPicture(0)   =   "NFe_Completa.frx":21B3A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "txtCodTransporte"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Frame6"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cboTransporte"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Volumes"
            TabPicture(1)   =   "NFe_Completa.frx":21B56
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label13"
            Tab(1).Control(1)=   "Label12"
            Tab(1).Control(2)=   "Label10"
            Tab(1).Control(3)=   "Label17"
            Tab(1).Control(4)=   "Label18"
            Tab(1).Control(5)=   "Label11"
            Tab(1).Control(6)=   "txtVolPesoLiquido"
            Tab(1).Control(7)=   "txtVolNumeracao"
            Tab(1).Control(8)=   "txtVolMarca"
            Tab(1).Control(9)=   "txtVolEspecie"
            Tab(1).Control(10)=   "txtVolQuant"
            Tab(1).Control(11)=   "txtVolPesoBruto"
            Tab(1).ControlCount=   12
            TabCaption(2)   =   "Reboques / Outros"
            TabPicture(2)   =   "NFe_Completa.frx":21B72
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame1"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Retenção do ICMS"
            TabPicture(3)   =   "NFe_Completa.frx":21B8E
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            Begin VB.ComboBox cboTransporte 
               Height          =   315
               Left            =   120
               TabIndex        =   112
               Top             =   660
               Width           =   7695
            End
            Begin VB.Frame Frame1 
               Caption         =   "Identificação"
               Height          =   1095
               Left            =   -74880
               TabIndex        =   126
               Top             =   420
               Width           =   12375
               Begin VB.TextBox txtPlacaReboque 
                  Height          =   315
                  Left            =   180
                  MaxLength       =   8
                  TabIndex        =   129
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.TextBox txtUFReboque 
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   2
                  TabIndex        =   128
                  Top             =   600
                  Width           =   465
               End
               Begin VB.TextBox txtRNTCReboque 
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   2
                  TabIndex        =   127
                  Top             =   600
                  Width           =   7305
               End
               Begin VB.Label Label49 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Placa"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   131
                  Top             =   360
                  Width           =   405
               End
               Begin VB.Label Label48 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   1500
                  TabIndex        =   130
                  Top             =   360
                  Width           =   210
               End
            End
            Begin VB.TextBox txtVolPesoBruto 
               Height          =   315
               Left            =   -70020
               MaxLength       =   50
               TabIndex        =   125
               Top             =   780
               Width           =   1545
            End
            Begin VB.TextBox txtVolQuant 
               Height          =   315
               Left            =   -74880
               MaxLength       =   50
               TabIndex        =   124
               Top             =   780
               Width           =   825
            End
            Begin VB.TextBox txtVolEspecie 
               Height          =   315
               Left            =   -74040
               MaxLength       =   50
               TabIndex        =   123
               Top             =   780
               Width           =   1245
            End
            Begin VB.TextBox txtVolMarca 
               Height          =   315
               Left            =   -72780
               MaxLength       =   50
               TabIndex        =   122
               Top             =   780
               Width           =   1665
            End
            Begin VB.TextBox txtVolNumeracao 
               Height          =   315
               Left            =   -71100
               MaxLength       =   50
               TabIndex        =   121
               Top             =   780
               Width           =   1065
            End
            Begin VB.TextBox txtVolPesoLiquido 
               Height          =   315
               Left            =   -68460
               MaxLength       =   50
               TabIndex        =   120
               Top             =   780
               Width           =   1545
            End
            Begin VB.Frame Frame6 
               Caption         =   "Veículo"
               Height          =   1095
               Left            =   120
               TabIndex        =   113
               Top             =   1080
               Width           =   12375
               Begin VB.TextBox txtTransRNTC 
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   2
                  TabIndex        =   116
                  Top             =   600
                  Width           =   7305
               End
               Begin VB.TextBox txtPlacaUF 
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   2
                  TabIndex        =   115
                  Top             =   600
                  Width           =   465
               End
               Begin VB.TextBox txtPlaca 
                  Height          =   315
                  Left            =   180
                  MaxLength       =   8
                  TabIndex        =   114
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label Label46 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RNTC"
                  Height          =   195
                  Left            =   2040
                  TabIndex        =   119
                  Top             =   360
                  Width           =   450
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   1500
                  TabIndex        =   118
                  Top             =   360
                  Width           =   210
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Placa"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   117
                  Top             =   360
                  Width           =   405
               End
            End
            Begin VB.TextBox txtCodTransporte 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7200
               MaxLength       =   50
               TabIndex        =   111
               Top             =   360
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Espécie"
               Height          =   195
               Left            =   -74040
               TabIndex        =   138
               Top             =   540
               Width           =   570
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso liquido"
               Height          =   195
               Left            =   -68460
               TabIndex        =   137
               Top             =   540
               Width           =   855
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso bruto"
               Height          =   195
               Left            =   -70020
               TabIndex        =   136
               Top             =   540
               Width           =   765
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde Vol."
               Height          =   195
               Left            =   -74880
               TabIndex        =   135
               Top             =   540
               Width           =   660
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Marca"
               Height          =   195
               Left            =   -72780
               TabIndex        =   134
               Top             =   540
               Width           =   450
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Numeração"
               Height          =   195
               Left            =   -71100
               TabIndex        =   133
               Top             =   540
               Width           =   825
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transportadora"
               Height          =   195
               Left            =   120
               TabIndex        =   132
               Top             =   420
               Width           =   1080
            End
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador de Pagamento:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   193
            Top             =   420
            Width           =   1785
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pagamento:"
            Height          =   195
            Left            =   -71640
            TabIndex        =   184
            Top             =   420
            Width           =   1560
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Frete"
            Height          =   195
            Left            =   -74880
            TabIndex        =   182
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave"
            Height          =   195
            Left            =   -74760
            TabIndex        =   143
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formato Impressão DANFE"
            Height          =   195
            Left            =   -74880
            TabIndex        =   142
            Top             =   480
            Width           =   1920
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Emissão NFe"
            Height          =   195
            Left            =   -72420
            TabIndex        =   141
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador Forma Pagto"
            Height          =   195
            Left            =   -74820
            TabIndex        =   140
            Top             =   420
            Width           =   1605
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade do Frete"
            Height          =   195
            Left            =   -74880
            TabIndex        =   139
            Top             =   420
            Width           =   1455
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   615
         Left            =   14700
         TabIndex        =   0
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":21BAA
         PICN            =   "NFe_Completa.frx":21BC6
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
         Left            =   14700
         TabIndex        =   47
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "NFe_Completa.frx":23958
         PICN            =   "NFe_Completa.frx":23974
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridNotas 
         Height          =   5940
         Left            =   -74880
         TabIndex        =   158
         Top             =   420
         Width           =   16395
         _ExtentX        =   28919
         _ExtentY        =   10478
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedidos 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   180
         Top             =   420
         Width           =   16395
         _ExtentX        =   28919
         _ExtentY        =   11668
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCopiarChave 
         Height          =   315
         Left            =   -65400
         TabIndex        =   192
         Top             =   6420
         Width           =   1155
         _ExtentX        =   2037
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
         MICON           =   "NFe_Completa.frx":25706
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
         Left            =   -69300
         TabIndex        =   194
         Top             =   6420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&DANFe"
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
         MICON           =   "NFe_Completa.frx":25722
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdTransmitir 
         Height          =   315
         Left            =   -73800
         TabIndex        =   195
         Top             =   6420
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Transmitir"
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
         MICON           =   "NFe_Completa.frx":2573E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarNota 
         Height          =   315
         Left            =   -72720
         TabIndex        =   196
         Top             =   6420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Cancelar NFe"
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
         MICON           =   "NFe_Completa.frx":2575A
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
         Height          =   315
         Left            =   -71520
         TabIndex        =   197
         Top             =   6420
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         MICON           =   "NFe_Completa.frx":25776
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdDuplicar 
         Height          =   315
         Left            =   -68400
         TabIndex        =   198
         Top             =   6420
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Duplicar"
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
         MICON           =   "NFe_Completa.frx":25792
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
         Height          =   315
         Left            =   -70500
         TabIndex        =   199
         Top             =   6420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Inutilizar NFe"
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
         MICON           =   "NFe_Completa.frx":257AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCartaCorrecao 
         Height          =   315
         Left            =   -66600
         TabIndex        =   200
         Top             =   6420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Carta Correção"
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
         MICON           =   "NFe_Completa.frx":257CA
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
         Left            =   -74880
         TabIndex        =   214
         Top             =   6420
         Width           =   1035
         _ExtentX        =   1826
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
         MICON           =   "NFe_Completa.frx":257E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEspelho 
         Height          =   315
         Left            =   -67440
         TabIndex        =   226
         Top             =   6420
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Espelho"
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
         MICON           =   "NFe_Completa.frx":25802
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConverterNFe 
         Height          =   315
         Left            =   -74880
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   7080
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Converter NFe"
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
         MICON           =   "NFe_Completa.frx":2581E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEnviarXML 
         Height          =   315
         Left            =   -64200
         TabIndex        =   237
         Top             =   6420
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "NFe_Completa.frx":2583A
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
         Left            =   -63060
         TabIndex        =   238
         Top             =   6420
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "NFe_Completa.frx":25856
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblQuantPedidos 
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
         Left            =   -58740
         TabIndex        =   181
         Top             =   7140
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   16605
      TabIndex        =   66
      Top             =   60
      Width           =   16635
      Begin VB.TextBox txtCodPedido 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   15300
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "Cód do Pedido"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCodNota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   14040
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Cód da Nota"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label MostraStatus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   13320
         TabIndex        =   225
         Top             =   240
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   540
         Picture         =   "NFe_Completa.frx":25872
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NFe - Nota Fiscal Eletrônica"
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
         Left            =   1500
         TabIndex        =   69
         Top             =   180
         Width           =   4140
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   65
      Top             =   9615
      Width           =   18210
      _ExtentX        =   32120
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25638
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18:44"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "NFe_Completa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moCombo As cComboHelper
Public TbNotas As New ADODB.Recordset
Public TbConsulta As New ADODB.Recordset
Public TbNotaPedido As New ADODB.Recordset
Dim Tb As New ADODB.Recordset
Dim Titulo, Book As Variant, NomeTabela
Dim TbProduto As New ADODB.Recordset
Private iRow As Long, iCol As Long, xCancelada As Boolean
Dim vCodCliente As String
Dim rEmpresa As ADODB.Recordset

Dim vTotalLinha As Currency
Dim vAliqLinha As Double
Dim vValorIcmsLinha As Currency
Dim vCodNota As Integer
Dim vSerieNota As Integer
Dim vTipoEdicaoNFe As String
Dim vTipoEdicaoNFeNFe As String
Dim vPossuiErro As Boolean

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Dim sSQL As String
Dim r As ADODB.Recordset
Dim vTipoCRT As Integer
Dim bSupressChkEvents As Boolean
Dim vRegimeTributario As Integer
Dim vIPICompoeDIFAL As Integer
Dim printSQL As String
Dim TipoSelecaoConsulta As String

Dim vEAN As String
Dim vInfAdd As String
Dim vDescricao As String
Dim vUnid_medida As String
Dim vCFOP As String
Dim vNCM As String
Dim vICMSCST As String
Dim vICMSAliq As String
Dim vpRedBC As String
Dim vModBC As String
Dim vPMVAST As String
Dim vPICMSST As String
Dim vPRedBCST As String
Dim vPISCST As String
Dim vPISALIQ As String
Dim vCOFINSCST As String
Dim vCOFINSALIQ As String
Dim vIPICST As String
Dim vIPIALIQ As String
Dim vCEST As String
Dim vTipoProduto As String

Public vAliqUFDest As Double
Public vAliqUFInter As Double
Public vUFEmpresa As String
Public vUFDest As String

'transportadora
Dim vTranspCNPJ As String
Dim vTranspEnd As String
Dim vTranspCidade As String
Dim vTranspUF As String
Dim vTranspIE As String

'parcelas e duplicatas
Dim vVencimento As Date
Dim vNumParc As Integer
Dim arrayParc() As Currency
Private Sub Calcular_Prazo()
If txtNumParcDup.Text = "" Then txtIntervaloDup.Text = "1": Exit Sub
If txtIntervaloDup.Text = "" Then txtIntervaloDup.Text = "0": Exit Sub
If mskEmissao.Text = "" Then Exit Sub

Dim vDataInicialCerta As Date

vDataInicialCerta = Format(mskEmissao.Text, "dd/mm/yy")

If txtIntervaloDup.Text = "30" Then
    mskInicioDup.Text = Format(DateAdd("m", Val(1), vDataInicialCerta), "dd/mm/yy")
Else
    mskInicioDup.Text = Format(DateAdd("d", Val(txtIntervaloDup.Text), vDataInicialCerta), "dd/mm/yy")
End If
End Sub






Private Sub AtualizarGrid_Itens()
AtualizarTotaisNota
End Sub

Private Sub RecalcularItensNota()
If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

Dim rItens       As ADODB.Recordset
Dim vItem        As Integer
Dim vValProd     As Currency
Dim sCST         As String
Dim dblPICMS     As Double
Dim dblPRedBC    As Double
Dim sModBC       As String
Dim sItemIPICST  As String
Dim dblPIPI      As Double
Dim dblPPIS      As Double
Dim dblPCOFINS   As Double
Dim dblPMVA      As Double
Dim dblPICMSST   As Double
Dim dblPRedBCST  As Double
Dim vValorIPI    As Currency
Dim curBaseICMS  As Currency
Dim curVICMS     As Currency
Dim curBasePISCOFINS As Currency
Dim curVPIS      As Currency
Dim curVCOFINS   As Currency
Dim curIPIvBC    As Currency
Dim dblIPIpGravar As Double
Dim curVBCST     As Currency
Dim curVICMSST   As Currency
Dim sCurModBC    As String
Dim sIPIcEnq     As String
Dim bSimples     As Boolean
Dim bDevolucao   As Boolean
Dim sUpd         As String

bSimples   = (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5)
bDevolucao = (Left(cboFinalidade.Text, 1) = "4")

sSQL = "SELECT ITEM, " & _
       "ValorUnitarioComercializacao * QuantidadeComercial AS vProd, " & _
       "CST, pICMS, pRedBC, modBC, " & _
       "IPICST, IPIpIPI, " & _
       "PISCST, PISpPIS, " & _
       "COFINSCST, cofinspcofins, " & _
       "pMVAST, pICMSST, pRedBCST " & _
       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen rItens, sSQL

Do While Not rItens.EOF
    vItem        = rItens("ITEM")
    vValProd     = CCur(rItens("vProd"))
    sCST         = Right(Format(rItens("CST"), "@"), 3)
    dblPICMS     = CDbl(IIf(IsNull(rItens("pICMS")),    0, rItens("pICMS")))
    dblPRedBC    = CDbl(IIf(IsNull(rItens("pRedBC")),   0, rItens("pRedBC")))
    sModBC       = Trim(IIf(IsNull(rItens("modBC")),    "", rItens("modBC")))
    sItemIPICST  = Trim(IIf(IsNull(rItens("IPICST")),   "", rItens("IPICST")))
    dblPIPI      = CDbl(IIf(IsNull(rItens("IPIpIPI")),  0, rItens("IPIpIPI")))
    dblPPIS      = CDbl(IIf(IsNull(rItens("PISpPIS")),  0, rItens("PISpPIS")))
    dblPCOFINS   = CDbl(IIf(IsNull(rItens("cofinspcofins")), 0, rItens("cofinspcofins")))
    dblPMVA      = CDbl(IIf(IsNull(rItens("pMVAST")),   0, rItens("pMVAST")))
    dblPICMSST   = CDbl(IIf(IsNull(rItens("pICMSST")),  0, rItens("pICMSST")))
    dblPRedBCST  = CDbl(IIf(IsNull(rItens("pRedBCST")), 0, rItens("pRedBCST")))

    ' IPI antecipado (necessario para base do ICMS com consumidor final)
    If bSimples And Not bDevolucao Then
        vValorIPI = 0
    Else
        vValorIPI = CCur(Format(vValProd * dblPIPI / 100, "0.00"))
    End If

    ' ICMS
    If bSimples And Not bDevolucao Then
        If sCST = "101" Or sCST = "201" Then
            curBaseICMS = vValProd
            If Left(cboConsumidorFinal.Text, 1) = "1" Then curBaseICMS = curBaseICMS + vValorIPI
            curVICMS = CCur(Format(curBaseICMS * dblPICMS / 100, "0.00"))
        Else
            curBaseICMS = 0
            curVICMS = 0
        End If
    Else
        If dblPRedBC > 0 Then
            curBaseICMS = CCur(vValProd * (1 - dblPRedBC / 100))
        Else
            curBaseICMS = vValProd
        End If
        If Left(cboConsumidorFinal.Text, 1) = "1" Then curBaseICMS = curBaseICMS + vValorIPI
        curVICMS = CCur(Format(curBaseICMS * dblPICMS / 100, "0.00"))
    End If

    ' modBC: vazio se nao houver base
    If curBaseICMS = 0 Then
        sCurModBC = ""
    Else
        sCurModBC = IIf(sModBC = "" Or sModBC = "0", "3", sModBC)
    End If

    ' PIS / COFINS
    If bSimples And Not bDevolucao Then
        curBasePISCOFINS = 0
        curVPIS   = 0
        curVCOFINS = 0
    Else
        curBasePISCOFINS = vValProd - curVICMS
        If curBasePISCOFINS < 0 Then curBasePISCOFINS = 0
        curVPIS    = CCur(Format(curBasePISCOFINS * dblPPIS   / 100, "0.00"))
        curVCOFINS = CCur(Format(curBasePISCOFINS * dblPCOFINS / 100, "0.00"))
    End If

    ' IPI gravar
    If bSimples And Not bDevolucao Then
        sIPIcEnq     = "999"
        curIPIvBC    = 0
        dblIPIpGravar = 0
        vValorIPI    = 0
    Else
        If sItemIPICST = "99" Or sItemIPICST = "53" Or sItemIPICST = "52" Or sItemIPICST = "50" Then
            sIPIcEnq = "999"
        Else
            sIPIcEnq = ""
        End If
        curIPIvBC    = vValProd
        dblIPIpGravar = dblPIPI
    End If

    ' ICMS-ST
    If chkICMSST.Value = 1 Then
        If bSimples And Not bDevolucao Then
            curVBCST   = 0
            curVICMSST = 0
        Else
            curVBCST = (vValProd + vValorIPI) * (1 + dblPMVA / 100)
            If dblPRedBCST > 0 Then curVBCST = curVBCST * (1 - dblPRedBCST / 100)
            curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
            If curVICMSST < 0 Then curVICMSST = 0
        End If
    Else
        curVBCST   = 0
        curVICMSST = 0
    End If

    sUpd = "UPDATE NotaFiscalItens SET " & _
           "modBC = '" & sCurModBC & "', " & _
           "vBC = " & FSQL(curBaseICMS, 2) & ", " & _
           "vICMS = " & FSQL(curVICMS, 2) & ", " & _
           "PISvBC = " & FSQL(curBasePISCOFINS, 2) & ", " & _
           "PISvPIS = " & FSQL(curVPIS, 2) & ", " & _
           "COFINSvBC = " & FSQL(curBasePISCOFINS, 2) & ", " & _
           "cofinsvcofins = " & FSQL(curVCOFINS, 2) & ", " & _
           "IPIcEnq = '" & sIPIcEnq & "', " & _
           "IPIvBC = " & FSQL(curIPIvBC, 2) & ", " & _
           "IPIpIPI = " & FSQL(dblIPIpGravar, 4) & ", " & _
           "IPIvIPI = " & FSQL(vValorIPI, 2) & ", " & _
           "vBCST = " & FSQL(curVBCST, 2) & ", " & _
           "vICMSST = " & FSQL(curVICMSST, 2) & " " & _
           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & vItem
    dbData.Execute sUpd

    rItens.MoveNext
Loop

Exibir_Itens
AtualizarTotaisNota
End Sub

Private Sub CalcularICMSInterItensGERAL()
If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

Dim vPICMSInter  As Double
Dim vPICMSUFDest As Double
Dim vPFCPUFDest  As Double
Dim vTipoCalc    As Integer
Dim vFCPBase     As Integer
Dim rDifal       As ADODB.Recordset

If cboDestOperacao.Text = "2 - Operação Interestadual" Then
    If cboConsumidorFinal.Text = "1 - SIM" Then

        ' 1. Aliquota interestadual (origem x destino)
        sSQL = "SELECT AliquotaInterestadual FROM TribMatrizInterestadual WHERE UF_Origem = '" & vUFEmpresa & "' AND UF_Destino = '" & vUFDest & "'"
        Set rDifal = dbData.OpenRecordset(sSQL)
        If rDifal.EOF Then
            MsgBox "Alíquota interestadual não encontrada: " & vUFEmpresa & " -> " & vUFDest, vbExclamation
            Exit Sub
        End If
        vPICMSInter = rDifal("AliquotaInterestadual")

        ' 2. Regras do estado de destino (vigente)
        sSQL = "SELECT TOP 1 AliquotaInterna, AliquotaFCP, TipoCalculo, FCPCompoeBase FROM TribRegraDifalUF " & _
               "WHERE UF_Destino = '" & vUFDest & "' AND DataInicioVigencia <= GETDATE() " & _
               "AND (DataFimVigencia IS NULL OR DataFimVigencia >= GETDATE()) " & _
               "ORDER BY DataInicioVigencia DESC"
        Set rDifal = dbData.OpenRecordset(sSQL)
        If rDifal.EOF Then
            MsgBox "Regra DIFAL não encontrada para: " & vUFDest, vbExclamation
            Exit Sub
        End If
        vPICMSUFDest = rDifal("AliquotaInterna")
        vPFCPUFDest = rDifal("AliquotaFCP")
        vTipoCalc = rDifal("TipoCalculo")
        vFCPBase = rDifal("FCPCompoeBase")

        ' 3. Gravar aliquotas e zerar campos nos itens
        sSQL = "UPDATE NotaFiscalItens SET pICMSInter = " & FSQL(vPICMSInter, 2) & ", pICMSUFDest = " & FSQL(vPICMSUFDest, 2) & ", pFCPUFDest = " & FSQL(vPFCPUFDest, 2) & ", pICMSInterPart = 100, vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL

        ' 4. Calcular vBCUFDest (base dupla ou simples)
        If vTipoCalc = 2 Then
            If vFCPBase = 1 Then
                sSQL = "UPDATE NotaFiscalItens SET " & _
                       "vBCUFDest    = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - (" & FSQL(vPICMSUFDest, 2) & " + " & FSQL(vPFCPUFDest, 2) & ") / 100), " & _
                       "vBCFCPUFDest = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - (" & FSQL(vPICMSUFDest, 2) & " + " & FSQL(vPFCPUFDest, 2) & ") / 100) " & _
                       "WHERE CodigoNota = " & Val(txtCodNota.Text)
            Else
                sSQL = "UPDATE NotaFiscalItens SET " & _
                       "vBCUFDest    = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - " & FSQL(vPICMSUFDest, 2) & " / 100), " & _
                       "vBCFCPUFDest = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - " & FSQL(vPICMSUFDest, 2) & " / 100) " & _
                       "WHERE CodigoNota = " & Val(txtCodNota.Text)
            End If
        Else
            sSQL = "UPDATE NotaFiscalItens SET " & _
                   "vBCUFDest    = vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100), " & _
                   "vBCFCPUFDest = vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100) " & _
                   "WHERE CodigoNota = " & Val(txtCodNota.Text)
        End If
        SQLExecuta sSQL

        ' 5. Calcular DIFAL e FCP por item
        sSQL = "UPDATE NotaFiscalItens SET " & _
               "vICMSUFDest = vBCUFDest    * (" & FSQL(vPICMSUFDest, 2) & " - " & FSQL(vPICMSInter, 2) & ") / 100, " & _
               "vFCPUFDest  = vBCFCPUFDest * " & FSQL(vPFCPUFDest, 2) & " / 100 " & _
               "WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL

        ' 6. Totalizar no cabecalho da nota
        sSQL = "UPDATE NotaFiscal SET " & _
               "vICMSUFDest = (SELECT ISNULL(SUM(vICMSUFDest), 0) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & "), " & _
               "vFCPUFDest  = (SELECT ISNULL(SUM(vFCPUFDest),  0) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & "), " & _
               "vICMSUFRemet = 0 " & _
               "WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL

    Else
        sSQL = "UPDATE NotaFiscalItens SET vBCUFDest = 0, vBCFCPUFDest = 0, pFCPUFDest = 0, pICMSUFDest = 0, pICMSInter = 0, pICMSInterPart = 0, vFCPUFDest = 0, vICMSUFRemet = 0, vICMSUFDest = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL
        sSQL = "UPDATE NotaFiscal SET vFCPUFDest = 0, vICMSUFDest = 0, vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL
    End If
Else
    sSQL = "UPDATE NotaFiscalItens SET vBCUFDest = 0, vBCFCPUFDest = 0, pFCPUFDest = 0, pICMSUFDest = 0, pICMSInter = 0, pICMSInterPart = 0, vFCPUFDest = 0, vICMSUFRemet = 0, vICMSUFDest = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    sSQL = "UPDATE NotaFiscal SET vFCPUFDest = 0, vICMSUFDest = 0, vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End If
End Sub
Private Sub CalcularICMSInterItens()
If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

Dim vPICMSInter  As Double
Dim vPICMSUFDest As Double
Dim vPFCPUFDest  As Double
Dim vTipoCalc    As Integer
Dim vFCPBase     As Integer
Dim rDifal       As ADODB.Recordset

If cboDestOperacao.Text = "2 - Operação Interestadual" Then
    If cboConsumidorFinal.Text = "1 - SIM" Then

        ' 1. Aliquota interestadual (origem x destino)
        sSQL = "SELECT AliquotaInterestadual FROM TribMatrizInterestadual WHERE UF_Origem = '" & vUFEmpresa & "' AND UF_Destino = '" & vUFDest & "'"
        Set rDifal = dbData.OpenRecordset(sSQL)
        If rDifal.EOF Then
            MsgBox "Alíquota interestadual não encontrada: " & vUFEmpresa & " -> " & vUFDest, vbExclamation
            Exit Sub
        End If
        vPICMSInter = rDifal("AliquotaInterestadual")

        ' 2. Regras do estado de destino (vigente)
        sSQL = "SELECT TOP 1 AliquotaInterna, AliquotaFCP, TipoCalculo, FCPCompoeBase FROM TribRegraDifalUF " & _
               "WHERE UF_Destino = '" & vUFDest & "' AND DataInicioVigencia <= GETDATE() " & _
               "AND (DataFimVigencia IS NULL OR DataFimVigencia >= GETDATE()) " & _
               "ORDER BY DataInicioVigencia DESC"
        Set rDifal = dbData.OpenRecordset(sSQL)
        If rDifal.EOF Then
            MsgBox "Regra DIFAL não encontrada para: " & vUFDest, vbExclamation
            Exit Sub
        End If
        vPICMSUFDest = rDifal("AliquotaInterna")
        vPFCPUFDest = rDifal("AliquotaFCP")
        vTipoCalc = rDifal("TipoCalculo")
        vFCPBase = rDifal("FCPCompoeBase")

        ' 3. Gravar aliquotas e zerar campos nos itens
        sSQL = "UPDATE NotaFiscalItens SET pICMSInter = " & FSQL(vPICMSInter, 2) & ", pICMSUFDest = " & FSQL(vPICMSUFDest, 2) & ", pFCPUFDest = " & FSQL(vPFCPUFDest, 2) & ", pICMSInterPart = 100, vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL

        ' 4. Calcular vBCUFDest (base dupla ou simples)
        If vTipoCalc = 2 Then
            If vFCPBase = 1 Then
                sSQL = "UPDATE NotaFiscalItens SET " & _
                       "vBCUFDest    = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - (" & FSQL(vPICMSUFDest, 2) & " + " & FSQL(vPFCPUFDest, 2) & ") / 100), " & _
                       "vBCFCPUFDest = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - (" & FSQL(vPICMSUFDest, 2) & " + " & FSQL(vPFCPUFDest, 2) & ") / 100) " & _
                       "WHERE CodigoNota = " & Val(txtCodNota.Text)
            Else
                sSQL = "UPDATE NotaFiscalItens SET " & _
                       "vBCUFDest    = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - " & FSQL(vPICMSUFDest, 2) & " / 100), " & _
                       "vBCFCPUFDest = (vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100)) / (1 - " & FSQL(vPICMSUFDest, 2) & " / 100) " & _
                       "WHERE CodigoNota = " & Val(txtCodNota.Text)
            End If
        Else
            sSQL = "UPDATE NotaFiscalItens SET " & _
                   "vBCUFDest    = vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100), " & _
                   "vBCFCPUFDest = vBC - (vBC * " & FSQL(vPICMSInter, 2) & " / 100) " & _
                   "WHERE CodigoNota = " & Val(txtCodNota.Text)
        End If
        SQLExecuta sSQL

        ' 5. Calcular DIFAL e FCP por item
        sSQL = "UPDATE NotaFiscalItens SET " & _
               "vICMSUFDest = vBCUFDest    * (" & FSQL(vPICMSUFDest, 2) & " - " & FSQL(vPICMSInter, 2) & ") / 100, " & _
               "vFCPUFDest  = vBCFCPUFDest * " & FSQL(vPFCPUFDest, 2) & " / 100 " & _
               "WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL

        ' 6. Totalizar no cabecalho da nota
        sSQL = "UPDATE NotaFiscal SET " & _
               "vICMSUFDest = (SELECT ISNULL(SUM(vICMSUFDest), 0) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & "), " & _
               "vFCPUFDest  = (SELECT ISNULL(SUM(vFCPUFDest),  0) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & "), " & _
               "vICMSUFRemet = 0 " & _
               "WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL

    Else
        sSQL = "UPDATE NotaFiscalItens SET vBCUFDest = 0, vBCFCPUFDest = 0, pFCPUFDest = 0, pICMSUFDest = 0, pICMSInter = 0, pICMSInterPart = 0, vFCPUFDest = 0, vICMSUFRemet = 0, vICMSUFDest = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL
        sSQL = "UPDATE NotaFiscal SET vFCPUFDest = 0, vICMSUFDest = 0, vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta sSQL
    End If
Else
    sSQL = "UPDATE NotaFiscalItens SET vBCUFDest = 0, vBCFCPUFDest = 0, pFCPUFDest = 0, pICMSUFDest = 0, pICMSInter = 0, pICMSInterPart = 0, vFCPUFDest = 0, vICMSUFRemet = 0, vICMSUFDest = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    sSQL = "UPDATE NotaFiscal SET vFCPUFDest = 0, vICMSUFDest = 0, vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End If
End Sub

Private Sub CalcularTotalProdutos()
If (GridNotas.TextMatrix(GridNotas.Row, 1)) = "" Then Exit Sub
vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))

Dim vTotalProdutoItens As Currency
Dim vTotalProdutoNota As Currency
'Dim vTotalProdutoItens As Currency
'(ValorUnitarioComercializacao * QuantidadeComercial)
'itens
'sSQL = "SELECT SUM(ValorTotalBruto) as ValorProdutosItens FROM NotaFiscalItens WHERE CodigoNota = " & Val(vCodNota)
sSQL = "SELECT SUM(ValorUnitarioComercializacao * QuantidadeComercial) as ValorProdutosItens FROM NotaFiscalItens WHERE CodigoNota = " & Val(vCodNota)
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
      vTotalProdutoItens = ValidateNull(r("ValorProdutosItens"))
End If

'nota
sSQL = "SELECT ValorProdutos FROM NotaFiscal WHERE CodigoNota = " & Val(vCodNota)
Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

If Not r.BOF Then
      vTotalProdutoNota = ValidateNull(r("ValorProdutos"))
End If

If vTotalProdutoItens <> vTotalProdutoNota Then
    sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotalProdutoItens, 2) & " WHERE CodigoNota = " & Val(vCodNota)
    dbData.Execute sSQL
End If
End Sub

Private Sub CorrecoesBasicasNFe()
Dim rCliente As ADODB.Recordset
Dim rNotaFiscal As ADODB.Recordset
Dim rNotaFiscalItens As ADODB.Recordset
Dim rEmpresa As ADODB.Recordset

'consultar o emitente
sSQL = "SELECT estado, CRT FROM empresa"
Set rEmpresa = dbData.OpenRecordset(sSQL)

Dim vUFEmpresa As String
Dim vCRTEmpresa As String

If Not rEmpresa.EOF Then
    vUFEmpresa = rEmpresa!Estado
    vCRTEmpresa = rEmpresa!CRT
End If

'Dim IdNFProd As Long
IdNFProd = GridNotas.TextMatrix(GridNotas.Row, 1)

'consultar o cpf/cnpj na nota
sSQL = "SELECT CodigoNota, CodigoCorrentista, IdentificadorDestino FROM NotaFiscal WHERE CodigoNota  = " & IdNFProd
Set rNotaFiscal = dbData.OpenRecordset(sSQL)

Dim vCodCliente As String

If Not rNotaFiscal.EOF Then
    vCodCliente = rNotaFiscal!CodigoCorrentista
End If

'consultar o cpf/cnpj do cliente
sSQL = "SELECT CODIGO, Nome, Endereco, Numero, Bairro, CEP, Cidade, Estado, CPF, CodigoIBGE, IE, TipoContribuinte, Tipo " & _
       "FROM cliente WHERE CODIGO  = " & vCodCliente
Set rCliente = dbData.OpenRecordset(sSQL)

'validação do cpf/cnpj do cliente
Dim vCPF As String
vCPF = ""

If Not rCliente.EOF Then
    vCPF = RetirarMascaras(rCliente!CPF)
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
                        MsgBox "CNPJ Informado não é válido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                        If ShowMsg("Deseja inserir o CNPJ na NFe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            vCPF = InputBox("Informe o CNPJ do cliente:", "EMISSÃO DE NFe", "")
                            If Not Vazio(vCPF) Then
                                If Len(vCPF) = 11 Then
                                    If Validar_CPF(vCPF) = False Then
                                        MsgBox "CPF Informado não é válido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                        GoTo CPFDigitadoErrado
                                    Else
                                        vCPF = Format(vCPF, "000\.000\.000\-00")
                                    End If
                                ElseIf Len(vCPF) = 14 Then
                                    If Validar_CNPJ(vCPF) = False Then
                                        MsgBox "CNPJ Informado não é válido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
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
                        MsgBox "CPF Informado não é válido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                        If ShowMsg("Deseja inserir o CPF na NFe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            vCPF = InputBox("Informe o CPF do cliente:", "EMISSÃO DE NFe", "")
                            If Not Vazio(vCPF) Then
                                If Len(vCPF) = 11 Then
                                    If Validar_CPF(vCPF) = False Then
                                        MsgBox "CPF Informado não é válido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
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

dbData.Execute "UPDATE cliente SET CPF = '" & vCPF & "' WHERE CODIGO  = " & vCodCliente
'FIM validação do cpf/cnpj do cliente

If vCPF = "" Then Exit Sub

If Not rCliente.EOF Then
    If rCliente!Tipo = "FÍSICA" And rCliente!TipoContribuinte = "1" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!" & Chr(13) & "Verifique o TIPO DE PESSOA e o TIPO DE CONTRIBUINTE no cadastro desse cliente!", vbCritical, "Online Commerce":  Exit Sub
    If rCliente!Tipo = "JURÍDICA" And rCliente!TipoContribuinte = "9" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!" & Chr(13) & "Verifique o TIPO DE PESSOA e o TIPO DE CONTRIBUINTE no cadastro desse cliente!", vbCritical, "Online Commerce":  Exit Sub
    If rCliente!Tipo = "RURAL" And rCliente!TipoContribuinte = "9" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!" & Chr(13) & "Verifique o TIPO DE PESSOA e o TIPO DE CONTRIBUINTE no cadastro desse cliente!", vbCritical, "Online Commerce":  Exit Sub
    If rCliente!Tipo = "RURAL" And rCliente!TipoContribuinte = "2" Then MsgBox "Tipo de contribuinte incompatível com o tipo de cadastro fiscal!" & Chr(13) & "Verifique o TIPO DE PESSOA e o TIPO DE CONTRIBUINTE no cadastro desse cliente!", vbCritical, "Online Commerce":  Exit Sub

    vCPF = RetirarMascaras(rCliente!CPF)
    
    'nome
    'Endereco
    'Numero
    'bairro
    'CEP
    'Cidade
    'Estado
    'CPF
    'CodigoIBGE
    'IE
    'TipoContribuinte
End If

'Destino do cliente         'DESATIVEI PQ DEU ERRO NO MAYERCK
'If rEmpresa!Estado = rCliente!Estado And rNotaFiscal!IdentificadorDestino = "2" Then
'    dbData.Execute "UPDATE NotaFiscal SET IdentificadorDestino = '1'  WHERE CodigoNota  = " & IdNFProd     'mayerck deu erro
'ElseIf rEmpresa!Estado <> rCliente!Estado And rNotaFiscal!IdentificadorDestino = "1" Then
'    dbData.Execute "UPDATE NotaFiscal SET IdentificadorDestino = '2'  WHERE CodigoNota  = " & IdNFProd
'End If

'consultar o destino e tipo de nota emitir e para onde vai
sSQL = "SELECT CodigoNota, IdentificadorDestino, TipoDocumento, FinalidadeEmissaoNFe, ConsumidorFinal FROM NotaFiscal WHERE CodigoNota  = " & IdNFProd
Set rNotaFiscal = dbData.OpenRecordset(sSQL)

    Dim vEmpresaForaEstado As Boolean
    vEmpresaForaEstado = False
    
    If Not rNotaFiscal.EOF Then
        'If vCRTEmpresa = 1 Then     'empresa simples nacional
            If rNotaFiscal!TipoDocumento = "1" And rNotaFiscal!FinalidadeEmissaoNFe = "1 - NFe NORMAL" Then
                If rNotaFiscal!IdentificadorDestino = "2" Then
                    vEmpresaForaEstado = True
                Else
                    vEmpresaForaEstado = False
                End If
            End If
        'End If                      'empresa simples nacional FIM
    End If


'correção cfop dos produtos da nota
    sSQL = "SELECT CodigoNota, CodigoProduto, CFOP, CST FROM NotaFiscalItens WHERE CodigoNota  = " & IdNFProd
    Set rNotaFiscalItens = dbData.OpenRecordset(sSQL)
    
    Dim vCodProdConsulta As Long
    vCodProdConsulta = "0"

If vEmpresaForaEstado = True Then
     For i = 1 To rNotaFiscalItens.RecordCount
        vCodProdConsulta = rNotaFiscalItens!CodigoProduto
         If vCRTEmpresa = 1 Then     'empresa simples nacional
            If rNotaFiscalItens!CFOP <> Empty Then
                If rNotaFiscalItens!CFOP = "5102" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6102', CST = '102' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                ElseIf rNotaFiscalItens!CFOP = "5405" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6403', CST = '500' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                End If
            Else
                dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6102', CST = '102' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
            End If
        Else                        'empresa lucro presumido ou lucro real
            If rNotaFiscalItens!CFOP <> Empty Then
                If rNotaFiscalItens!CFOP = "5102" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6102' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                ElseIf rNotaFiscalItens!CFOP = "5405" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                End If
            Else
                dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
            End If
        End If                      'empresa simples nacional FIM

     rNotaFiscalItens.MoveNext
     Next
Else
     For i = 1 To rNotaFiscalItens.RecordCount
        vCodProdConsulta = rNotaFiscalItens!CodigoProduto
         If vCRTEmpresa = 1 Then     'empresa simples nacional
            If rNotaFiscalItens!CFOP <> Empty Then
                If rNotaFiscalItens!CFOP = "6102" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5102', CST = '102' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                ElseIf rNotaFiscalItens!CFOP = "6405" Or rNotaFiscalItens!CFOP = "6403" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5403', CST = '500' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                End If
            Else
                dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5102', CST = '102' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
            End If
        Else                        'empresa lucro presumido ou lucro real
            If rNotaFiscalItens!CFOP <> Empty Then
                If rNotaFiscalItens!CFOP = "6102" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5102' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                ElseIf rNotaFiscalItens!CFOP = "5405" Or rNotaFiscalItens!CFOP = "6403" Then
                    dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
                End If
            Else
                dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = " & vCodProdConsulta & " And CodigoNota = " & IdNFProd
            End If
        End If                      'empresa simples nacional FIM

     rNotaFiscalItens.MoveNext
     Next
End If

'imposto inter estadual
If vEmpresaForaEstado = True Then
    If Len(vCPF) = 14 And rCliente!TipoContribuinte = "1" And rCliente!IE <> Empty Then
        dbData.Execute "UPDATE NotaFiscal SET ConsumidorFinal = '0', vFCPUFDest = 0, vICMSUFDest = 0, vICMSUFRemet = 0  WHERE CodigoNota  = " & IdNFProd
        dbData.Execute "UPDATE NotaFiscalItens SET vBCUFDest = 0, vBCFCPUFDest = 0, pFCPUFDest = 0, pICMSUFDest = 0, pICMSInter = 0, pICMSInterPart = 0, vFCPUFDest = 0, vICMSUFRemet = 0, vICMSUFDest = 0 WHERE CodigoNota = " & IdNFProd
    End If
End If

End Sub

Private Sub LimparObjestosNotaOutros()
cboModFrete.Text = ""
txtCodTransporte.Text = ""
cboTransporte.Text = ""
txtPlaca.Text = ""
txtPlacaUF.Text = ""
txtTransRNTC.Text = ""
txtVolQuant.Text = ""
txtVolEspecie.Text = ""
txtVolMarca.Text = ""
txtVolNumeracao.Text = ""
txtVolPesoBruto.Text = ""
txtVolPesoLiquido.Text = ""
txtPlacaReboque.Text = ""
txtUFReboque.Text = ""
txtRNTCReboque.Text = ""
cboIndicadorPagamento.Text = ""
cboFormaPgto.Text = ""
txtNumFatura.Text = ""
txtSubtotalFatura.Text = ""
txtDescFatura.Text = ""
txtTotalFatura.Text = ""
txtNumDup.Text = ""
txtTotalDup.Text = ""
txtNumParcDup.Text = ""
txtIntervaloDup.Text = ""
mskInicioDup.Mask = ""
mskInicioDup.Text = ""
txtValorParcDup.Text = ""
txtInfComple.Text = ""
txtInfAdicionais.Text = ""
cboFormatoDANFe.Text = ""
cboTipoEmissao.Text = ""
txtChaveReferenciada.Text = ""
End Sub

Private Sub LimparObjetosDestinatario()
txtCodCliente.Text = ""
cboTipoDest.Text = ""
cboCliente.Text = ""
cboTipoContribuinte.Text = ""
cboConsumidorFinal.Text = ""
txtAliqUFDest.Text = ""
End Sub

Private Sub LimparObjetosNota()
txtCodNota.Text = "0"
txtSerie.Text = "0"
txtCodPedido.Text = "0"
txtNumNota.Text = ""
cboTipoNota.Text = ""
cboFinalidade.Text = ""
cboDestOperacao.Text = ""
cboNatureza.Text = ""
txtNatureza.Text = ""
mskEmissao.Mask = ""
mskEmissao.Text = ""
mskSaida.Mask = ""
mskSaida.Text = ""
mskHora.Mask = ""
mskHora.Text = ""
MostraStatus.Caption = ""
frmDuplicata.Visible = False
End Sub


Private Sub LimparObjetosNotaTotais()
txtBaseICMS.Text = FormatNumber(0, 2)
txtBaseICMSST.Text = FormatNumber(0, 2)
txtTotaldosProdutos.Text = FormatNumber(0, 2)
txtValorFrete.Text = FormatNumber(0, 2)
txtValorICMS.Text = FormatNumber(0, 2)
txtValorICMSST.Text = FormatNumber(0, 2)
txtValorIPI.Text = FormatNumber(0, 2)
txtValorDesconto.Text = FormatNumber(0, 2)
txtValorSeguro.Text = FormatNumber(0, 2)
txtValorOutrasDespesas.Text = FormatNumber(0, 2)
txtTotaldaNota.Text = FormatNumber(0, 2)
End Sub

Private Sub MostrarCorrecao()
If vTipoEdicaoNFe <> "Novo" Then
    If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
    
    vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))
    
    sSQL = "SELECT CodigoCartaCorrecao, CodigoNota, Data, SeqCorrecao, TextoCorrecao, NumeroProtocolo, DataHoraProcotolo, (CASE WHEN Enviada = 1 THEN 'ENVIADO' ELSE 'NÃO ENVIADO' END) as vStatusCorrecao FROM NFeCartaCorrecao WHERE (CodigoNota = " & vCodNota & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGridCorrecao r
    If r.State <> 0 Then r.Close
End If
End Sub

Private Sub FormatarGridCorrecao(rTabela As ADODB.Recordset)
Dim j As Integer

With Grid_Correcao
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 500
   .ColWidth(2) = 700
   .ColWidth(3) = 900
   .ColWidth(4) = 700
   .ColWidth(5) = 5000
   .ColWidth(6) = 900
   .ColWidth(7) = 900
   .ColWidth(8) = 2000

   .TextMatrix(0, 1) = "CÓD"
   .TextMatrix(0, 2) = "NOTA"
   .TextMatrix(0, 3) = "DATA"
   .TextMatrix(0, 4) = "EVENTO"
   .TextMatrix(0, 5) = "TEXTO"
   .TextMatrix(0, 6) = "PROTOCOLO"
   .TextMatrix(0, 7) = "DATA"
   .TextMatrix(0, 8) = "STATUS"

   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next i

   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
           'CodigoCartaCorrecao, CodigoNota, Data, SeqCorrecao, TextoCorrecao, NumeroProtocolo, DataHoraProcotolo, Enviada

         .TextMatrix(.rows - 1, 1) = rTabela("CodigoCartaCorrecao")
         .TextMatrix(.rows - 1, 2) = rTabela("CodigoNota")
         .TextMatrix(.rows - 1, 3) = Format(rTabela("Data"), ocDATA2)
         .TextMatrix(.rows - 1, 4) = rTabela("SeqCorrecao")
         .TextMatrix(.rows - 1, 5) = rTabela("TextoCorrecao")
         .TextMatrix(.rows - 1, 6) = rTabela("NumeroProtocolo")
         .TextMatrix(.rows - 1, 7) = Format(rTabela("DataHoraProcotolo"), ocDATA2)
         .TextMatrix(.rows - 1, 8) = rTabela("vStatusCorrecao")
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .rows = .rows - 1
   .Redraw = True
End With
End Sub
Private Sub MostrarValorNota()
Dim varValorFrete As Currency
Dim varValorICMS As Currency
Dim varValorICMSST As Currency
Dim varValorIPI As Currency
Dim varValorDesconto As Currency
Dim varValorSeguro As Currency
Dim varValorOutrasDespesas As Currency
Dim varTotalProdutos As Currency
Dim varTotalNota As Currency

If txtTotaldosProdutos.Text = "" Then varTotalProdutos = 0 Else varTotalProdutos = txtTotaldosProdutos.Text
If txtValorFrete.Text = "" Then varValorFrete = 0 Else varValorFrete = txtValorFrete.Text
If txtValorICMS.Text = "" Then varValorICMS = 0 Else varValorICMS = txtValorICMS.Text
If txtValorICMSST.Text = "" Then varValorICMSST = 0 Else varValorICMSST = txtValorICMSST.Text
If txtValorIPI.Text = "" Then varValorIPI = 0 Else varValorIPI = txtValorIPI.Text
If txtValorDesconto.Text = "" Then varValorDesconto = 0 Else varValorDesconto = txtValorDesconto.Text
If txtValorSeguro.Text = "" Then varValorSeguro = 0 Else varValorSeguro = txtValorSeguro.Text
If txtValorOutrasDespesas.Text = "" Then varValorOutrasDespesas = 0 Else varValorOutrasDespesas = txtValorOutrasDespesas.Text

'varTotalNota = varTotalProdutos + varValorFrete + varValorICMS + varValorIPI + varValorSeguro + varValorOutrasDespesas
varTotalNota = varTotalProdutos
varTotalNota = varTotalNota + txtValorFrete + txtValorIPI + varValorSeguro + varValorOutrasDespesas + varValorICMSST
varTotalNota = varTotalNota - varValorDesconto
txtTotaldaNota = FormatNumber(varTotalNota, 2)

'Parte de faturas
txtNumFatura.Text = txtCodNota.Text
txtSubtotalFatura.Text = FormatNumber(varTotalProdutos, 2)
txtDescFatura.Text = FormatNumber(varValorDesconto, 2)
txtTotalFatura.Text = FormatNumber(varTotalNota, 2)
txtNumDup.Text = txtCodNota.Text
txtTotalDup.Text = FormatNumber(varTotalNota, 2)
txtNumParcDup.Text = "1"
txtIntervaloDup.Text = "30"
If IsDate(mskEmissao) Then mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy") Else: mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy")
End Sub

Private Sub AtualizarTotaisNota()
Dim rTotais       As ADODB.Recordset
Dim varICMSST     As Double
Dim varBaseICMSST As Double
Dim varProdutos   As Double
Dim varFrete      As Double
Dim varSeguro     As Double
Dim varOutras     As Double
Dim varDesconto   As Double
Dim varIPI        As Double
Dim varICMS       As Double
Dim varBaseICMS   As Double
Dim varICMSUFDest As Double
Dim varFCPUFDest  As Double
Dim varPIS        As Double
Dim varCOFINS     As Double
Dim varNota       As Double

If txtCodNota.Text = "" Then Exit Sub

' Todos os totais vem dos itens
sSQL = "SELECT " & _
       "ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS ValorProdutos, " & _
       "ISNULL(SUM(ValorFrete),   0) AS ValorFrete,   " & _
       "ISNULL(SUM(ValorSeguro),  0) AS ValorSeguro,  " & _
       "ISNULL(SUM(ValorOutros),  0) AS ValorOutros,  " & _
       "ISNULL(SUM(ValorDesconto),0) AS ValorDesconto," & _
       "ISNULL(SUM(IPIvIPI),      0) AS ValorIPI,     " & _
       "ISNULL(SUM(vICMS),        0) AS ValorICMS,    " & _
       "ISNULL(SUM(vBC),          0) AS BaseICMS,     " & _
       "ISNULL(SUM(vBCST),        0) AS BaseICMSST,   " & _
       "ISNULL(SUM(vICMSST),      0) AS ValorICMSST,  " & _
       "ISNULL(SUM(PISvPIS),      0) AS ValorPIS,     " & _
       "ISNULL(SUM(cofinsvcofins),0) AS ValorCOFINS,  " & _
       "ISNULL(SUM(vICMSUFDest),  0) AS vICMSUFDest,  " & _
       "ISNULL(SUM(vFCPUFDest),   0) AS vFCPUFDest    " & _
       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
Set rTotais = dbData.OpenRecordset(sSQL)

If Not rTotais.EOF Then
    varProdutos = ValidateNull(rTotais("ValorProdutos"))
    varFrete = ValidateNull(rTotais("ValorFrete"))
    varSeguro = ValidateNull(rTotais("ValorSeguro"))
    varOutras = ValidateNull(rTotais("ValorOutros"))
    varDesconto = ValidateNull(rTotais("ValorDesconto"))
    varIPI = ValidateNull(rTotais("ValorIPI"))
    varICMS = ValidateNull(rTotais("ValorICMS"))
    varBaseICMS = ValidateNull(rTotais("BaseICMS"))
    varICMSUFDest = ValidateNull(rTotais("vICMSUFDest"))
    varFCPUFDest = ValidateNull(rTotais("vFCPUFDest"))
    varICMSST = ValidateNull(rTotais("ValorICMSST"))
    varBaseICMSST = ValidateNull(rTotais("BaseICMSST"))
    varPIS = ValidateNull(rTotais("ValorPIS"))
    varCOFINS = ValidateNull(rTotais("ValorCOFINS"))
End If

varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST - varDesconto

' Preencher textboxes dos totais
txtTotaldosProdutos.Text = FormatNumber(varProdutos, 2)
txtValorFrete.Text = FormatNumber(varFrete, 2)
txtValorSeguro.Text = FormatNumber(varSeguro, 2)
txtValorOutrasDespesas.Text = FormatNumber(varOutras, 2)
txtValorDesconto.Text = FormatNumber(varDesconto, 2)
txtValorIPI.Text = FormatNumber(varIPI, 2)
txtValorICMS.Text = FormatNumber(varICMS, 2)
txtBaseICMS.Text = FormatNumber(varBaseICMS, 2)
txtValorICMSST.Text = FormatNumber(varICMSST, 2)
txtBaseICMSST.Text = FormatNumber(varBaseICMSST, 2)
txtTotaldaNota.Text = FormatNumber(varNota, 2)

' Persistir na tabela NotaFiscal
sSQL = "UPDATE NotaFiscal SET " & _
       "ValorProdutos       = " & FSQL(varProdutos, 2) & ", " & _
       "ValorFrete          = " & FSQL(varFrete, 2) & ", " & _
       "ValorSeguro         = " & FSQL(varSeguro, 2) & ", " & _
       "ValorOutrasDespesas = " & FSQL(varOutras, 2) & ", " & _
       "ValorDesconto       = " & FSQL(varDesconto, 2) & ", " & _
       "ValorIPI            = " & FSQL(varIPI, 2) & ", " & _
       "ValorICMS           = " & FSQL(varICMS, 2) & ", " & _
       "BaseICMS            = " & FSQL(varBaseICMS, 2) & ", " & _
       "ValorICMSST         = " & FSQL(varICMSST, 2) & ", " & _
       "BaseICMSST          = " & FSQL(varBaseICMSST, 2) & ", " & _
       "ValorPIS            = " & FSQL(varPIS, 2) & ", " & _
       "ValorCOFINS         = " & FSQL(varCOFINS, 2) & ", " & _
       "vICMSUFDest         = " & FSQL(varICMSUFDest, 2) & ", " & _
       "vFCPUFDest          = " & FSQL(varFCPUFDest, 2) & ", " & _
       "ValorNota           = " & FSQL(varNota, 2) & _
       " WHERE CodigoNota = " & Val(txtCodNota.Text)
SQLExecuta sSQL

' Fatura
txtNumFatura.Text = txtCodNota.Text
txtSubtotalFatura.Text = FormatNumber(varProdutos, 2)
txtDescFatura.Text = FormatNumber(varDesconto, 2)
txtTotalFatura.Text = FormatNumber(varNota, 2)
txtNumDup.Text = txtCodNota.Text
txtTotalDup.Text = FormatNumber(varNota, 2)
txtNumParcDup.Text = "1"
txtIntervaloDup.Text = "30"
If IsDate(mskEmissao) Then mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy") Else mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy")
End Sub

Private Sub CalcularICMSInterNota()
If txtCodNota.Text = "" Then Exit Sub

If GridNotasItens.rows <= 1 Then Exit Sub

sSQL = "SELECT SUM(vICMSUFDest) as ValorICMSDest FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

Dim vValorTotalIcmsDest As Currency
If Not r.BOF Then
      vValorTotalIcmsDest = Format(ValidateNull(r("ValorICMSDest")), ocMONEY)
End If

sSQL = "UPDATE NotaFiscal SET vFCPUFDest = 0, vICMSUFDest = " & FSQL(vValorTotalIcmsDest, 2) & ", vICMSUFRemet = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
SQLExecuta sSQL
End Sub

Private Sub CalcularDesconto()
If txtCodNota.Text = "" Then Exit Sub

'If GridNotasItens.Rows <= 1 Then Exit Sub

sSQL = "SELECT SUM(Valordesconto) as ValorDesc FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
      txtValorDesconto = Format(ValidateNull(r("ValorDesc")), ocMONEY)
End If
End Sub

Private Sub DistribuirDesconto()
Dim vTotalDesc      As Currency
Dim vTotalSubtotal  As Currency
Dim vResto          As Currency
Dim rTot            As ADODB.Recordset
Dim rResto          As ADODB.Recordset

If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

If txtValorDesconto.Text <> "0" And txtValorDesconto.Text <> "" Then
    vTotalDesc = txtValorDesconto.Text
Else
    vTotalDesc = 0
End If

If vTotalDesc = 0 Then Exit Sub

' Busca subtotal total dos itens
sSQL = "SELECT ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS TotalSubtotal " & _
       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
Set rTot = dbData.OpenRecordset(sSQL)
If rTot.EOF Then Exit Sub
vTotalSubtotal = CCur(rTot("TotalSubtotal"))
If vTotalSubtotal = 0 Then Exit Sub

' Valida: desconto nao pode exceder o subtotal total dos produtos
If vTotalDesc > vTotalSubtotal Then
    ShowMsg "O desconto total (" & FormatNumber(vTotalDesc, 2) & ") nao pode ser maior que o subtotal dos produtos (" & FormatNumber(vTotalSubtotal, 2) & ").", vbExclamation
    txtValorDesconto.Text = FormatNumber(vTotalSubtotal, 2)
    Exit Sub
End If

' Distribui proporcionalmente ao subtotal de cada item
sSQL = "UPDATE NotaFiscalItens SET " & _
       "TipoDesconto = 1, " & _
       "Desconto     = ROUND(" & FSQL(vTotalDesc, 2) & " * (ValorUnitarioComercializacao * QuantidadeComercial) / " & FSQL(vTotalSubtotal, 2) & ", 2), " & _
       "ValorDesconto = ROUND(" & FSQL(vTotalDesc, 2) & " * (ValorUnitarioComercializacao * QuantidadeComercial) / " & FSQL(vTotalSubtotal, 2) & ", 2) " & _
       "WHERE CodigoNota = " & Val(txtCodNota.Text)
SQLExecuta sSQL

' Ajusta o resto do arredondamento no item com maior subtotal (tem mais margem)
sSQL = "SELECT " & FSQL(vTotalDesc, 2) & " - ISNULL(SUM(ValorDesconto), 0) AS Resto " & _
       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
Set rResto = dbData.OpenRecordset(sSQL)
If Not rResto.EOF Then vResto = CCur(rResto("Resto"))

If vResto <> 0 Then
    sSQL = "UPDATE TOP(1) NotaFiscalItens SET " & _
           "ValorDesconto = ValorDesconto + " & FSQL(vResto, 2) & ", " & _
           "Desconto      = ValorDesconto + " & FSQL(vResto, 2) & " " & _
           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _
           "AND (ValorUnitarioComercializacao * QuantidadeComercial) = " & _
           "(SELECT MAX(ValorUnitarioComercializacao * QuantidadeComercial) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ")"
    SQLExecuta sSQL
End If
End Sub

Private Sub MostrarValorItens()
If txtCodNota.Text = "" Then Exit Sub

'If GridNotasItens.Rows <= 1 Then Exit Sub

sSQL = "UPDATE NotaFiscalItens SET " & _
       "ValorTotalBruto = ((ValorUnitarioComercializacao * QuantidadeComercial) - ValorDesconto),  vBC = ((ValorUnitarioComercializacao * QuantidadeComercial) - ValorDesconto) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
SQLExecuta sSQL

Exibir_Itens
End Sub

Private Sub MostrarValorProdutos()
If txtCodNota.Text = "" Then Exit Sub
'sSQL = "SELECT SUM(ValorTotalBruto) as ValorProdutos FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
sSQL = "SELECT SUM(ValorUnitarioComercializacao * QuantidadeComercial) as ValorProdutos FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

If Not r.BOF Then
      txtTotaldosProdutos = FormatNumber(ValidateNull(r("ValorProdutos")), 2)
End If
End Sub

Private Sub CalcularIPI()
If txtCodNota.Text = "" Then Exit Sub

If GridNotasItens.rows <= 1 Then Exit Sub

sSQL = "SELECT SUM(IPIvIPI) as ValorIPI FROM NotaFiscalItens WHERE (IPIvIPI <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
      txtValorIPI = Format(ValidateNull(r("ValorIPI")), ocMONEY)
End If

End Sub

Private Sub MostrarValorBaseICMS()
If txtCodNota.Text = "" Then Exit Sub
Dim vBaseICMS As Currency
Dim vValorICMS As Currency

Dim varValorFrete As Currency
Dim varValorIPI As Currency
Dim varValorDesconto As Currency
Dim varValorSeguro As Currency
Dim varValorOutrasDespesas As Currency
Dim varValorProdutos As Currency

'If txtValorFrete.Text = "" Then varValorFrete = 0 Else varValorFrete = txtValorFrete.Text
'If txtValorIPI.Text = "" Then varValorIPI = 0 Else varValorIPI = txtValorIPI.Text
'If txtValorDesconto.Text = "" Then varValorDesconto = 0 Else varValorDesconto = txtValorDesconto.Text
'If txtValorSeguro.Text = "" Then varValorSeguro = 0 Else varValorSeguro = txtValorSeguro.Text
'If txtValorOutrasDespesas.Text = "" Then varValorOutrasDespesas = 0 Else varValorOutrasDespesas = txtValorOutrasDespesas.Text

'If GridNotasItens.Rows <= 1 Then Exit Sub

'frete
sSQL = "SELECT SUM(ValorFrete) as vValorFrete FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)
varValorFrete = ValidateNull(r("vValorFrete"))

'seguro
sSQL = "SELECT SUM(ValorSeguro) as vValorSeguro FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)
varValorSeguro = ValidateNull(r("vValorSeguro"))

'outras
sSQL = "SELECT SUM(ValorOutros) as vValorOutros FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)
varValorOutrasDespesas = ValidateNull(r("vValorOutros"))

'ipi
sSQL = "SELECT SUM(IPIvIPI) as vValorIPI FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)
varValorIPI = ValidateNull(r("vValorIPI"))

'Desconto
sSQL = "SELECT SUM(ValorDesconto) as vValorDesc FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)
varValorDesconto = ValidateNull(r("vValorDesc"))

'valor dos produtos
sSQL = "SELECT SUM(ValorUnitarioComercializacao * QuantidadeComercial) as vValorProdutos FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
      varValorProdutos = Format(ValidateNull(r("vValorProdutos")), ocMONEY)
End If

'sSQL = "SELECT SUM(ValorTotalBruto) as ValorBaseICMS FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
'Set r = dbData.OpenRecordset(sSQL)

vBaseICMS = varValorProdutos + varValorFrete + varValorIPI + varValorSeguro + varValorOutrasDespesas
vBaseICMS = vBaseICMS - varValorDesconto

'If Not r.BOF Then
txtBaseICMS.Text = Format(vBaseICMS, ocMONEY)
      'txtBaseICMS.Text = Format(ValidateNull(r("ValorBaseICMS")), ocMONEY)
'End If

sSQL = "SELECT SUM(vICMS) as ValorICMS FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
      txtValorICMS = Format(ValidateNull(r("ValorICMS")), ocMONEY)
End If

End Sub
Private Sub DistribuirSeguro()
Dim varQuantItens          As Integer
Dim vTotalSeguro           As Currency
Dim vValorSeguroIndividual As Currency
Dim vValorSeguroAjuste     As Currency

If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

sSQL = "SELECT codigonota FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

If Not Tb.EOF Then
    varQuantItens = Tb.RecordCount
Else
    varQuantItens = 0
End If

If txtValorSeguro.Text <> "0" And txtValorSeguro.Text <> "" Then
    vTotalSeguro = txtValorSeguro.Text
Else
    vTotalSeguro = 0
End If

If vTotalSeguro = 0 Or varQuantItens = 0 Then
    Exit Sub
Else
    vValorSeguroIndividual = CCur(Format(vTotalSeguro / varQuantItens, "0.00"))
    vValorSeguroAjuste = vTotalSeguro - (vValorSeguroIndividual * (varQuantItens - 1))
    
    sSQL = "UPDATE NotaFiscalItens SET ValorSeguro = " & FSQL(vValorSeguroIndividual, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorSeguro = " & FSQL(vValorSeguroAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End If
End Sub

Private Sub DistribuirOutros()
Dim varQuantItens          As Integer
Dim vTotalOutros           As Currency
Dim vValorOutrosIndividual As Currency
Dim vValorOutrosAjuste     As Currency

If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

sSQL = "SELECT codigonota FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

If Not Tb.EOF Then
    varQuantItens = Tb.RecordCount
Else
    varQuantItens = 0
End If

If txtValorOutrasDespesas.Text <> "0" And txtValorOutrasDespesas.Text <> "" Then
    vTotalOutros = txtValorOutrasDespesas.Text
Else
    vTotalOutros = 0
End If

If vTotalOutros = 0 Or varQuantItens = 0 Then
    Exit Sub
Else
    vValorOutrosIndividual = CCur(Format(vTotalOutros / varQuantItens, "0.00"))
    vValorOutrosAjuste = vTotalOutros - (vValorOutrosIndividual * (varQuantItens - 1))
    
    sSQL = "UPDATE NotaFiscalItens SET ValorOutros = " & FSQL(vValorOutrosIndividual, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorOutros = " & FSQL(vValorOutrosAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End If
End Sub
Private Sub AtualizarValorICMS()
sSQL = "SELECT pRedBC as AliqRedBC FROM NotaFiscalItens WHERE (pRedBC <> '0.0000') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then  'calculo de base de calculo redusida por incentivo de icms
    Dim vAliqReducao As Double
    vAliqReducao = r("AliqRedBC")
    
    'sSQL = "UPDATE NotaFiscalItens SET vBC = " & _
           "CASE CST WHEN 051 THEN 0 ELSE  (((ValorTotalBruto + ValorFrete + ValorSeguro + ValorOutros) * " & FSQL(vAliqReducao, 4) & ") / 100) END FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    
    sSQL = "Update NotaFiscalItens SET vBC = " & _
            "CASE CST WHEN 051 THEN 0 ELSE CAST(((ValorTotalBruto + ValorFrete + ValorSeguro + ValorOutros) * (1 - (" & FSQL(vAliqReducao, 4) & " / 100.0))) AS DECIMAL(18, 2)) End " & _
            "From NotaFiscalItens " & _
            "Where CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    
    sSQL = "UPDATE NotaFiscalItens SET " & _
           "vICMS = ((vBC * pICMS) / 100) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
Else
    sSQL = "UPDATE NotaFiscalItens SET vBC = " & _
           "CASE CST WHEN 051 THEN 0 ELSE (ValorTotalBruto + ValorFrete + ValorSeguro + ValorOutros) END FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    
    sSQL = "UPDATE NotaFiscalItens SET " & _
           "vICMS = ((vBC * pICMS) / 100) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End If

sSQL = "SELECT SUM(vICMS) as ValorICMS, SUM(vBC) as ValorBC FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

Dim vValorTotalICMS As Currency
Dim vValorBaseICMS As Currency

If Not r.BOF Then
      vValorTotalICMS = Format(ValidateNull(r("ValorICMS")), ocMONEY)
      vValorBaseICMS = Format(ValidateNull(r("ValorBC")), ocMONEY)
      txtValorICMS = Format(ValidateNull(r("ValorICMS")), ocMONEY)
      txtBaseICMS = Format(ValidateNull(r("ValorBC")), ocMONEY)
End If

sSQL = "UPDATE NotaFiscal SET  valorICMS = " & FSQL(vValorTotalICMS, 2) & ", BaseICMS = " & FSQL(vValorBaseICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
SQLExecuta sSQL


'CALCULO DE ICMS FUNCIONANDO NORMAL ANTES DE ACRESCENTAR O REDUÇÃO DE ICMS
'sSQL = "UPDATE NotaFiscalItens SET vBC = " & _
       "CASE CST WHEN 051 THEN 0 ELSE (ValorTotalBruto + ValorFrete + ValorSeguro + ValorOutros) END FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
'SQLExecuta sSQL

'sSQL = "UPDATE NotaFiscalItens SET " & _
       "vICMS = ((vBC * pICMS) / 100) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
'SQLExecuta sSQL

'sSQL = "SELECT SUM(vICMS) as ValorICMS, SUM(vBC) as ValorBC FROM NotaFiscalItens WHERE (vICMS <> '0.00') AND CodigoNota = " & Val(txtCodNota.Text)
'Set r = dbData.OpenRecordset(sSQL)

'Dim vValorTotalICMS As Currency
'Dim vValorBaseICMS As Currency

'If Not r.BOF Then
'      vValorTotalICMS = Format(ValidateNull(r("ValorICMS")), ocMONEY)
'      vValorBaseICMS = Format(ValidateNull(r("ValorBC")), ocMONEY)
'      txtValorICMS = Format(ValidateNull(r("ValorICMS")), ocMONEY)
'      txtBaseICMS = Format(ValidateNull(r("ValorBC")), ocMONEY)
'End If

'sSQL = "UPDATE NotaFiscal SET  valorICMS = " & FSQL(vValorTotalICMS, 2) & ", BaseICMS = " & FSQL(vValorBaseICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
'SQLExecuta sSQL
End Sub
Private Sub DistribuirFrete()
Dim varQuantItens       As Integer
Dim varValorTotalFrete  As Currency
Dim varValorDividoFrete As Currency
Dim varValorAjusteFrete As Currency

If txtCodNota.Text = "" Then Exit Sub
If GridNotasItens.rows <= 1 Then Exit Sub

sSQL = "SELECT codigonota FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

If Not Tb.EOF Then
    varQuantItens = Tb.RecordCount
Else
    varQuantItens = 0
End If

If txtValorFrete.Text <> "0" And txtValorFrete.Text <> "" Then
    varValorTotalFrete = txtValorFrete.Text
Else
    varValorTotalFrete = 0
End If

If varValorTotalFrete = 0 Or varQuantItens = 0 Then
    Exit Sub
Else
    varValorDividoFrete = CCur(Format(varValorTotalFrete / varQuantItens, "0.00"))
    varValorAjusteFrete = varValorTotalFrete - (varValorDividoFrete * (varQuantItens - 1))
    
    sSQL = "UPDATE NotaFiscalItens SET valorfrete = " & FSQL(varValorDividoFrete, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    sSQL = "UPDATE TOP(1) NotaFiscalItens SET valorfrete = " & FSQL(varValorAjusteFrete, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End If
End Sub
Private Sub Exibir_Cliente()
'If cboTipoDest.Text = "CLIENTE" Then
'    Dim TbClientes As New ADODB.Recordset
'    RsOpen TbClientes, "SELECT * FROM cliente WHERE codigo = " & Val(txtCodCliente.Text)
'    If Not TbClientes.EOF Then
'        cboCliente.Text = ValidateNull(TbClientes("nome"))
'    End If
'ElseIf cboTipoDest.Text = "FORNECEDOR" Then

'End If
End Sub

Private Sub Exibir_Duplicatas()
If txtCodNota.Text = "" Then Exit Sub

sSQL = "SELECT *, " & _
"(CASE WHEN CodigoFormaPagamento = 1  THEN 'Dinheiro' " & _
"WHEN CodigoFormaPagamento = 2 THEN 'Cheque' " & _
"WHEN CodigoFormaPagamento = 3 THEN 'Cartão de Crédito' " & _
"WHEN CodigoFormaPagamento = 4 THEN 'Cartão de Débito' " & _
"WHEN CodigoFormaPagamento = 5 THEN 'Crédito Loja' " & _
"WHEN CodigoFormaPagamento = 10 THEN 'Vale Alimentação' " & _
"WHEN CodigoFormaPagamento = 11 THEN 'Vale Refeição' " & _
"WHEN CodigoFormaPagamento = 12 THEN 'Vale Presente' " & _
"WHEN CodigoFormaPagamento = 13 THEN 'Vale Combustível' " & _
"WHEN CodigoFormaPagamento = 14 THEN 'Duplicata Mercantil' " & _
"WHEN CodigoFormaPagamento = 15 THEN 'Boleto Bancário' " & _
"WHEN CodigoFormaPagamento = 16 THEN 'Depósito Bancário' " & _
"WHEN CodigoFormaPagamento = 17 THEN 'PIX' " & _
"WHEN CodigoFormaPagamento = 18 THEN 'Transferência bancária' " & _
"WHEN CodigoFormaPagamento = 19 THEN 'Programa de fidelidade' " & _
"WHEN CodigoFormaPagamento = 90 THEN 'Sem pagamento' " & _
"WHEN CodigoFormaPagamento = 99 THEN 'Outros' " & _
"Else 'Dinheiro' END) As var_FormaPgto " & _
"FROM NotaFiscalParcelas " & _
"WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _
"ORDER BY Sequencia;"

RsOpen Tb, sSQL

FormatarGridDuplicatas Tb
End Sub
Private Sub Exibir_Itens()
If txtCodNota.Text = "" Then Exit Sub

sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, " & _
       "ValorUnitarioComercializacao, QuantidadeComercial, ValorTotalBruto, " & _
       "ValorFrete, ValorSeguro, ValorOutros, ValorDesconto, " & _
       "vBC, pICMS, vICMS, pRedBC, " & _
       "vBCST, pICMSST, vICMSST, pMVAST, " & _
       "IPICST, IPIpIPI, IPIvIPI " & _
       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

FormatarGridItensNota Tb
AplicarVisibilidadeGridItens
End Sub


Private Sub ExibirUltimasNfe()
Dim totalRegistros As Long

RsOpen TbConsulta, "SELECT top 25 CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal order by NumeroNota desc"

If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount
'lblQuantEnviada.Caption = Format(totalRegistros, "00")

LimparGridNotas
FormatarGridNotas TbConsulta
'lblTotalEnviada.Caption = Format(SomaGrid(GridNotas, 6), ocMONEY)

Dim soma As Currency
Dim contar As Integer
Dim i As Integer

'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Enviada" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantEnviada.Caption = Format(contar, "000")
lblTotalEnviada.Caption = Format(soma, ocMONEY)

'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Cancelada" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantCancelada.Caption = Format(contar, "000")
lblTotalCancelada.Caption = Format(soma, ocMONEY)

'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Em Digitação" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantNaoEnviada.Caption = Format(contar, "000")
lblTotalNaoEnviada.Caption = Format(soma, ocMONEY)


Exit Sub
Resume
End Sub

Private Sub GravarPedido()
flag = False

'On Error GoTo Err_Grava

Dim r As ADODB.Recordset
Dim totalRegistros As Long

'If txtCodPedido = "" Then Exit Sub

'preencher objetos da nota com o pedido
sSQL = "SELECT pedidos.*, cliente.codigo, cliente.nome as VarNome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE pedidos.cod_pedido = " & txtCodPedido & ";"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.BOF Then Mostrar_Pedido r

If r.State <> 0 Then r.Close
Set r = Nothing

If txtCodCliente.Text = "" Then MsgBox "O campo CLIENTE é obrigatório.", vbCritical, "Online Commerce": txtCodCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo código CFOP é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
'If txtCodObservacao.Text = "" Then MsgBox "O campo mensagem é obrigatório.", vbCritical, "Online Commerce": txtCodObservacao.SetFocus: Exit Sub

If txtCodPedido.Text = "0" Then

Else
    RsOpen TbNotas, "SELECT * FROM NotaFiscal"
    TbNotas.AddNew
End If

flag = True

Load_Data
TbNotas.Update
vgDb.CommitTrans

Load_Controls

TransformarPedidoemNFE  'TESTE

SomarProdutosNota

PreencherGridNotas

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False

'Clear_Controls
LimparObjetosProduto
End Sub



Private Sub LerDadosInserir()
'On Error GoTo erro
'    vgDb.BeginTrans
    If txtVolPesoBruto.Text = "" Then txtVolPesoBruto.Text = "0"
    If txtVolPesoLiquido.Text = "" Then txtVolPesoLiquido.Text = "0"
    TbNotas("CodigoNatureza") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("CodigoNota") = Format(txtCodNota.Text, "@")
    
    'TbNotas("SerieNF") = 0

    TbNotas("TipoCliente") = Format(cboTipoDest, "@")
    TbNotas("InformacoesComplementares") = Format(txtInfComple, "@")
    TbNotas("CFOP") = Format(cboNatureza.Text, "@")
    TbNotas("NaturezaOperacao") = Format(Left(txtNatureza, 59), "@")
    TbNotas("TipoDocumento") = IIf(IsNull(Format(Left(cboTipoNota.Text, 1), "@")) Or Vazio(Format(Left(cboTipoNota.Text, 1), "@")), 1, Format(Left(cboTipoNota.Text, 1), "@"))
    TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")

    TbNotas("DataEmissao") = Format(Date, "dd/mm/yyyy")
    TbNotas("DataSaida") = Format(Date, "dd/mm/yyyy")
    TbNotas("HoraSaida") = Format(mskHora, "@")
    
    TbNotas("FinalidadeEmissaoNFe") = Format(cboFinalidade, "@")
    TbNotas("CodigoObservacao") = IIf(IsNull(Format(txtCodObservacao, "@")) Or Vazio(Format(txtCodObservacao, "@")), 0, Format(txtCodObservacao, "@"))
    TbNotas("NumeroNota") = Format(txtNumNota, "@")
    TbNotas("cCodigoNota") = IIf(TbNotas("cCodigoNota") = 0, GeraCodigoNota, TbNotas("cCodigoNota"))

    TbNotas("CodigoCorrentista") = IIf(IsNull(Format(txtCodCliente, "@")) Or Vazio(Format(txtCodCliente, "@")), 0, Format(txtCodCliente, "@"))
    TbNotas("RazaoSocial") = Format(cboCliente, "@")
    TbNotas("IndicadorFormaPagamento") = Format(cboIndicadorPagamento.Text, "@")
    TbNotas("FormatoImpressaoDANFE") = Format(cboFormatoDANFe.Text, "@")
    TbNotas("FormatoEmissaoNFe") = Format(cboTipoEmissao.Text, "@")
    TbNotas("IdentificadorDestino") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("IndicadorIEDestinatario") = IIf(IsNull(Format(Left(cboTipoContribuinte.Text, 1), "@")) Or Vazio(Format(Left(cboTipoContribuinte.Text, 1), "@")), 1, Format(Left(cboTipoContribuinte.Text, 1), "@"))
    TbNotas("ConsumidorFinal") = IIf(IsNull(Format(Left(cboConsumidorFinal.Text, 1), "@")) Or Vazio(Format(Left(cboConsumidorFinal.Text, 1), "@")), 1, Format(Left(cboConsumidorFinal.Text, 1), "@"))
    TbNotas("ChavedeAcessoAdicional") = Format(txtChaveReferenciada.Text, "@")

    'tributos e valores
    TbNotas("BaseICMS") = IIf(IsNull(Format(txtBaseICMS, "@")) Or Vazio(Format(txtBaseICMS, "@")), 0, CDbl(Format(txtBaseICMS, "##0.00")))
    TbNotas("BaseICMSST") = IIf(IsNull(Format(txtBaseICMSST, "@")) Or Vazio(Format(txtBaseICMSST, "@")), 0, CDbl(Format(txtBaseICMSST, "##0.00")))
    TbNotas("ValorFrete") = IIf(Vazio(txtValorFrete), 0, CDbl(Format(txtValorFrete, "##0.00")))
    TbNotas("ValorSeguro") = IIf(IsNull(Format(txtValorSeguro, "@")) Or Vazio(Format(txtValorSeguro, "@")), 0, CDbl(Format(txtValorSeguro, "##0.00")))
    TbNotas("ValorOutrasDespesas") = IIf(IsNull(Format(txtValorOutrasDespesas, "@")) Or Vazio(Format(txtValorOutrasDespesas, "@")), 0, CDbl(Format(txtValorOutrasDespesas, "##0.00")))
    TbNotas("ValorICMS") = IIf(IsNull(Format(txtValorICMS, "@")) Or Vazio(Format(txtValorICMS, "@")), 0, CDbl(Format(txtValorICMS, "##0.000")))
    TbNotas("ValorICMSST") = IIf(IsNull(Format(txtValorICMSST, "@")) Or Vazio(Format(txtValorICMSST, "@")), 0, CDbl(Format(txtValorICMSST, "##0.00")))
    TbNotas("ValorIPI") = IIf(IsNull(Format(txtValorIPI, "@")) Or Vazio(Format(txtValorIPI, "@")), 0, CDbl(Format(txtValorIPI, "##0.000")))
'    TbNotas("ValorProdutos") = IIf(IsNull(Format(txtTotaldosProdutos, "@")) Or Vazio(Format(txtTotaldosProdutos, "@")), 0, CDbl(Format(txtTotaldosProdutos, ocPESO)))
    TbNotas("ValorDesconto") = IIf(IsNull(Format(txtValorDesconto, "@")) Or Vazio(Format(txtValorDesconto, "@")), 0, CDbl(Format(txtValorDesconto, "##0.00")))

'    TbNotas("valornota") = IIf(IsNull(Format(txtTotaldaNota, "@")) Or Vazio(Format(txtTotaldaNota, "@")), 0, CDbl(Format(txtTotaldaNota, ocPESO)))

    'TbNotas("BaseICMS") = " & FSQL(txtBaseICMS, 2) & "
    'TbNotas("BaseICMSST") = " & FSQL(txtBaseICMSST, 2) & "
    'TbNotas("ValorFrete") = Format(txtValorFrete.Text, "@")
    'TbNotas("ValorSeguro") = Format(txtValorSeguro.Text, "@")
    'TbNotas("ValorOutrasDespesas") = Format(txtValorOutrasDespesas.Text, "@")
    'TbNotas("ValorICMS") = Format(txtValorICMS.Text, "@")
    'TbNotas("ValorICMSST") = Format(txtValorICMSST.Text, "@")
    'TbNotas("ValorIPI") = Format(txtValorIPI.Text, "@")
    'TbNotas("ValorDesconto") = Format(txtValorDesconto.Text, "@")

    'transporte
    TbNotas("ModFrete") = IIf(IsNull(Format(Left(cboModFrete.Text, 1), "@")) Or Vazio(Format(Left(cboModFrete.Text, 1), "@")), 9, Format(Left(cboModFrete.Text, 1), "@"))
    TbNotas("TranspCodigo") = IIf(IsNull(Format(txtCodTransporte, "@")) Or Vazio(Format(txtCodTransporte, "@")), 0, Format(txtCodTransporte, "@"))
    TbNotas("TranspNome") = Format(cboTransporte, "@")
    TbNotas("TranspPlaca") = Format(txtPlaca, "@")
    TbNotas("TranspPlacaUF") = Format(txtPlacaUF, "@")
    TbNotas("VolumeQuantidade") = Format(txtVolQuant, "@")
    TbNotas("VolumeEspecie") = Format(txtVolEspecie, "@")
    TbNotas("VolumeMarca") = Format(txtVolMarca, "@")
    TbNotas("VolumeNumeracao") = Format(txtVolNumeracao, "@")
    TbNotas("VolumePesoBruto") = IIf(IsNull(Format(txtVolPesoBruto, "@")) Or Vazio(Format(txtVolPesoBruto, "@")), 0, CDbl(Format(txtVolPesoBruto, "##0.000")))
    TbNotas("VolumePesoLiquido") = IIf(IsNull(Format(txtVolPesoLiquido, "@")) Or Vazio(Format(txtVolPesoLiquido, "@")), 0, CDbl(Format(txtVolPesoLiquido, "##0.000")))

    TbNotas("SerieNF") = 2
    TbNotas("InscricaoEstadual") = 0
    TbNotas("Suframa") = 0
    TbNotas("CNPJ_CPF") = 0
    TbNotas("Logradouro") = 0
    TbNotas("numero") = 0
    TbNotas("CodigoIBGE") = 0
    TbNotas("Bairro") = 0
    TbNotas("Complemento") = 0
    TbNotas("Municipio") = 0
    TbNotas("UF") = 0
    TbNotas("CEP") = 0
    TbNotas("CODIGOPAIS") = 0
    TbNotas("PAIS") = 0
    TbNotas("TELEFONE") = 0
    'Exit Sub

'Resume

'erro:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
'    vgDb.RollbackTrans
'    Exit Sub
End Sub



Private Sub LimparVariaveisItens()
vEAN = ""
vInfAdd = ""
vUnid_medida = ""
vCFOP = ""
vNCM = ""
vICMSCST = ""
vICMSAliq = ""
vpRedBC = ""
vPISCST = ""
vPISALIQ = ""
vCOFINSCST = ""
vCOFINSALIQ = ""
vIPICST = ""
vIPIALIQ = ""
vCEST = ""
End Sub

Sub Load_Data_Duplicatas()
Dim seq As Integer

If txtCodNota.Text = "" Then Exit Sub

sSQL = "SELECT MAX(Item) r FROM NotaFiscalParcelas WHERE CodigoNota = " & Val(txtCodNota.Text)
seq = SQLExecutaRetorno(sSQL, "r", 0) + 1

'CodigoNota, Sequencia, Documento, CodigoFormaPagamento, Vencimento, ValorDocumento

'txtNumDup
'txtTotalDup
'txtNumParcDup
'txtIntervaloDup
'mskInicioDup
'txtValorParcDup
'cmdCalDuplic

Tb("CodigoNota") = Format(txtCodNota.Text, "@")
Tb("Sequencia") = seq
Tb("Documento") = Format(txtNumDup.Text, "@")
Tb("CodigoFormaPagamento") = IIf(IsNull(Format(Left(cboFormaPgto.Text, 2), "@")) Or Vazio(Format(Left(cboFormaPgto.Text, 2), "@")), 1, Format(Left(cboFormaPgto.Text, 2), "@"))
Tb("Vencimento") = IIf(Tb("Vencimento") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskInicioDup, "@"))
Tb("ValorDocumento") = CDbl(Format(txtValorParcDup, "@"))

End Sub
Sub Load_Data_Itens()
Dim seq As Integer
Dim vValorProdutos As Currency
Dim vPorcICMS As Double
Dim vValorICMS As Currency
Dim vPorcIPI As Double
Dim vValorIPI As Currency
Dim vValorPIS As Currency
Dim vValorCOFINS As Currency
Dim rDifalLDI    As ADODB.Recordset
Dim vPICMSInter  As Double
Dim vPICMSUFDest As Double
Dim vPFCPUFDest  As Double
Dim vTipoCalcLDI As Integer
Dim vFCPBaseLDI  As Integer
Dim vBaseItem    As Double
Dim vBCUFDestLDI As Double
Dim vICMSUFDestLDI As Double
Dim vFCPUFDestLDI  As Double
Dim vCSOSN As String
Dim curBasePISCOFINS As Currency
Dim curBaseICMS As Currency
Dim curVBCST As Currency
Dim curVICMSST As Currency
Dim dblMVAFinal As Double
Dim dblAliqInter As Double
Dim dblAliqInterna As Double

If txtCodNota.Text = "" Then Exit Sub
    sSQL = "SELECT MAX(Item) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    seq = SQLExecutaRetorno(sSQL, "r", 0) + 1
    Tb("CodigoNota") = Format(txtCodNota.Text, "@")
    Tb("Item") = seq
    Tb("CodigoProduto") = Format(txtCodProduto, "@")
    Tb("NomeProduto") = UCase(Format(cboDescricao, "@"))
    Tb("InformacoesAdicionaisProduto") = UCase(Format(vInfAdd, "@"))
    Tb("TipoProduto") = Format(vTipoProduto, "@")
    
    'tributos
    Tb("EAN") = Format(vEAN, "@")
    ' CFOP: converter 5xxx -> 6xxx se operacao interestadual
    Dim sCFOPFinal As String
    sCFOPFinal = Format(vCFOP, "@")
    If Left(cboDestOperacao.Text, 1) = "2" Then
        If Left(sCFOPFinal, 1) = "5" Then sCFOPFinal = "6" & Mid(sCFOPFinal, 2)
    End If
    Tb("CFOP") = sCFOPFinal
    Tb("NCM") = Format(vNCM, "@")
    Tb("UnidadeComercial") = UCase(Format(vUnid_medida, "@"))
    vValorProdutos = CDbl(Format(txtSubTotal, "@"))
    
    'ICMS
    ' CST/CSOSN: para Simples (1,2), derivar do CFOP original do cadastro
    Dim sCSTFinal As String
    If vRegimeTributario = 1 Or vRegimeTributario = 2 Then
        If Right(Format(vCFOP, "@"), 3) = "102" Then
            sCSTFinal = "102"
        ElseIf Right(Format(vCFOP, "@"), 3) = "405" Then
            sCSTFinal = "500"
        Else
            sCSTFinal = Right(Format(vICMSCST, "@"), 3)
        End If
    Else
        sCSTFinal = Right(Format(vICMSCST, "@"), 3)
    End If
    vICMSCST = sCSTFinal
    Tb("CST") = sCSTFinal
    Tb("modBC") = Format(IIf(vModBC = "", 3, vModBC), "@")
    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))
    If vpRedBC = "" Then Tb("pRedBC") = CDbl(Format(0, "@")) Else Tb("pRedBC") = CDbl(Format(vpRedBC, "@"))
    ' Calculo antecipado do IPI para usar na base do ICMS (consumidor final)
    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then
        vValorIPI = 0
    Else
        If vIPIALIQ = "" Then
            vValorIPI = 0
        Else
            vValorIPI = CCur(Format(CCur(vValorProdutos) * CDbl(vIPIALIQ) / 100, "0.00"))
        End If
    End If
    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then
        ' Simples Nacional venda normal: zera vBC/vICMS exceto CSOSN 101 e 201
        vCSOSN = Right(Format(vICMSCST, "@"), 3)
        If vCSOSN = "101" Or vCSOSN = "201" Then
            vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))
            curBaseICMS = CCur(txtSubTotal.Text)
            If Left(cboConsumidorFinal.Text, 1) = "1" Then curBaseICMS = curBaseICMS + vValorIPI
            Tb("vBC") = CDbl(Format(curBaseICMS, "0.00"))
            vValorICMS = CCur(Format(curBaseICMS * vPorcICMS / 100, "0.00"))
            Tb("vICMS") = CDbl(Format(vValorICMS, "@"))
        Else
            Tb("vBC") = CDbl(Format(0, "@"))
            Tb("vICMS") = CDbl(Format(0, "@"))
        End If
    Else
        ' Regime Normal ou devolucao: aplica reducao de BC + IPI se consumidor final
        If vpRedBC <> "" And CDbl(vpRedBC) > 0 Then
            curBaseICMS = CCur(txtSubTotal.Text) * (1 - CDbl(vpRedBC) / 100)
        Else
            curBaseICMS = CCur(txtSubTotal.Text)
        End If
        If Left(cboConsumidorFinal.Text, 1) = "1" Then curBaseICMS = curBaseICMS + vValorIPI
        Tb("vBC") = CDbl(Format(curBaseICMS, "0.00"))
        vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))
        vValorICMS = CCur(Format(curBaseICMS * vPorcICMS / 100, "0.00"))
        If vICMSAliq = "" Then Tb("vICMS") = CDbl(Format(0, "@")) Else Tb("vICMS") = CDbl(Format(vValorICMS, "@"))
    End If
    
    'PIS e COFINS
    If vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5 Then
        ' Simples Nacional / MEI: nao destaca PIS/COFINS por item
        Tb("PISCST") = Right(Format(vPISCST, "@"), 2)
        Tb("PISvBC") = CDbl(Format(0, "@"))
        Tb("PISpPIS") = CDbl(Format(0, "@"))
        Tb("PISvPIS") = CDbl(Format(0, "@"))
        Tb("PISqBCProd") = CDbl(Format(0, "@"))
        Tb("PISvAliqProd") = CDbl(Format(0, "@"))
        Tb("COFINSCST") = Right(Format(vCOFINSCST, "@"), 2)
        Tb("COFINSvBC") = CDbl(Format(0, "@"))
        Tb("cofinspcofins") = CDbl(Format(0, "@"))
        Tb("cofinsvcofins") = CDbl(Format(0, "@"))
        Tb("COFINSqBCProd") = CDbl(Format(0, "@"))
        Tb("COFINSvAliqProd") = CDbl(Format(0, "@"))
    Else
        ' Regime Normal: base = valor liquido - ICMS (Tese do Seculo STF RE 574.706)
        curBasePISCOFINS = CCur(vValorProdutos) - vValorICMS
        If curBasePISCOFINS < 0 Then curBasePISCOFINS = 0
        Tb("PISCST") = Right(Format(vPISCST, "@"), 2)
        Tb("PISvBC") = CDbl(Format(curBasePISCOFINS, "0.00"))
        If vPISALIQ = "" Then
            Tb("PISpPIS") = CDbl(Format(0, "@"))
            Tb("PISvPIS") = CDbl(Format(0, "@"))
        Else
            Tb("PISpPIS") = CDbl(Format(vPISALIQ, "@"))
            vValorPIS = CCur(Format(curBasePISCOFINS * CDbl(vPISALIQ) / 100, "0.00"))
            Tb("PISvPIS") = CDbl(Format(vValorPIS, "@"))
        End If
        Tb("PISqBCProd") = CDbl(Format(0, "@"))
        Tb("PISvAliqProd") = CDbl(Format(0, "@"))
        Tb("COFINSCST") = Right(Format(vCOFINSCST, "@"), 2)
        Tb("COFINSvBC") = CDbl(Format(curBasePISCOFINS, "0.00"))
        If vCOFINSALIQ = "" Then
            Tb("cofinspcofins") = CDbl(Format(0, "@"))
            Tb("cofinsvcofins") = CDbl(Format(0, "@"))
        Else
            Tb("cofinspcofins") = CDbl(Format(vCOFINSALIQ, "@"))
            vValorCOFINS = CCur(Format(curBasePISCOFINS * CDbl(vCOFINSALIQ) / 100, "0.00"))
            Tb("cofinsvcofins") = CDbl(Format(vValorCOFINS, "@"))
        End If
        Tb("COFINSqBCProd") = CDbl(Format(0, "@"))
        Tb("COFINSvAliqProd") = CDbl(Format(0, "@"))
    End If

    'IPI
    Tb("IPICST") = Format(vIPICST, "@")
    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then
        ' Simples Nacional venda normal: zera IPI
        Tb("IPIcEnq") = "999"
        Tb("IPIvBC") = CDbl(Format(0, "@"))
        Tb("IPIpIPI") = CDbl(Format(0, "@"))
        Tb("IPIvIPI") = CDbl(Format(0, "@"))
    Else
        ' Regime Normal ou devolucao: calcula IPI (vValorIPI ja calculado na secao ICMS)
        If vIPICST = "99" Or vIPICST = "53" Or vIPICST = "52" Or vIPICST = "50" Then
            Tb("IPIcEnq") = "999"
        Else
            Tb("IPIcEnq") = ""
        End If
        Tb("IPIvBC") = CDbl(Format(vValorProdutos, "0.00"))
        If vIPIALIQ = "" Then
            Tb("IPIpIPI") = CDbl(Format(0, "@"))
            Tb("IPIvIPI") = CDbl(Format(0, "@"))
        Else
            Tb("IPIpIPI") = CDbl(Format(vIPIALIQ, "@"))
            Tb("IPIvIPI") = CDbl(Format(vValorIPI, "@"))
        End If
    End If
    
    'Valores do item
    Tb("ValorUnitarioComercializacao") = CDbl(Format(txtValor, "@"))
    If txtQuant.Text <> "" Then Tb("QuantidadeComercial") = CDbl(Format(txtQuant, "@")) Else Tb("QuantidadeComercial") = Format(1, "@")
    Tb("ValorFrete") = CDbl(IIf(txtFrete.Text = "", 0, Format(txtFrete, "@")))
    Tb("ValorSeguro") = CDbl(IIf(txtSeguro.Text = "", 0, Format(txtSeguro, "@")))
    Tb("ValorOutros") = CDbl(IIf(txtOutrosItem.Text = "", 0, Format(txtOutrosItem, "@")))
    Tb("tipodesconto") = Format(1, "@")
    Tb("desconto") = CDbl(IIf(txtDesc.Text = "", 0, Format(txtDesc, "@")))
    Tb("Valordesconto") = CDbl(IIf(txtDesc.Text = "", 0, Format(txtDesc, "@")))
    Tb("ValorTotalBruto") = CDbl(Format(txtSubTotal, "@"))
    
    Tb("referencia") = Format(0, "@")

    'ICMS-ST
    If chkICMSST.Value = 1 Then
        ' chkICMSST marcado: calcula ou copia ST dependendo do regime e finalidade
        If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then
            ' Simples Nacional venda normal: zera ST (ST ja retido anteriormente)
            Tb("modBCST") = Format(0, "@")
            Tb("pMVAST") = CDbl(Format(0, "@"))
            Tb("pRedBCST") = CDbl(Format(0, "@"))
            Tb("vBCST") = CDbl(Format(0, "@"))
            Tb("pICMSST") = CDbl(Format(0, "@"))
            Tb("vICMSST") = CDbl(Format(0, "@"))
        Else
            ' Regime Normal ou devolucao: calcula ST via MVA
            dblAliqInterna = CDbl(IIf(vPICMSST = "" Or vPICMSST = "0,00", 0, vPICMSST))
            If vUFEmpresa <> vUFDest And vUFDest <> "" Then
                Set rDifalLDI = dbData.OpenRecordset("SELECT AliquotaInterestadual FROM TribMatrizInterestadual WHERE UF_Origem = '" & vUFEmpresa & "' AND UF_Destino = '" & vUFDest & "'")
                If Not rDifalLDI.EOF Then
                    dblAliqInter = CDbl(rDifalLDI("AliquotaInterestadual"))
                Else
                    dblAliqInter = CDbl(IIf(vAliqUFInter = 0, 0, vAliqUFInter))
                End If
                rDifalLDI.Close
            Else
                dblAliqInter = 0
            End If

            ' MVA: ajustado se interestadual, original se interna
            If vUFEmpresa <> vUFDest And vUFDest <> "" And dblAliqInterna > 0 Then
                Dim dblMVAOrig As Double
                dblMVAOrig = CDbl(IIf(vPMVAST = "" Or vPMVAST = "0,00", 0, vPMVAST))
                If (1 - dblAliqInterna / 100) <> 0 Then
                    dblMVAFinal = (((1 + dblMVAOrig / 100) * (1 - dblAliqInter / 100)) / (1 - dblAliqInterna / 100) - 1) * 100
                    dblMVAFinal = Int(dblMVAFinal * 100 + 0.5) / 100
                Else
                    dblMVAFinal = dblMVAOrig
                End If
            Else
                dblMVAFinal = CDbl(IIf(vPMVAST = "" Or vPMVAST = "0,00", 0, vPMVAST))
            End If

            ' Base do ST: (valor liquido + IPI) * (1 + MVA/100)
            curVBCST = (CCur(vValorProdutos) + vValorIPI) * (1 + dblMVAFinal / 100)

            ' Reducao da base ST se houver
            If vPRedBCST <> "" And CDbl(vPRedBCST) > 0 Then
                curVBCST = curVBCST * (1 - CDbl(vPRedBCST) / 100)
            End If

            ' vICMSST = (vBCST * aliq. interna) - ICMS proprio; nunca negativo
            curVICMSST = (curVBCST * dblAliqInterna / 100) - vValorICMS
            If curVICMSST < 0 Then curVICMSST = 0

            Tb("modBCST") = Format(4, "@")
            Tb("pMVAST") = CDbl(Format(dblMVAFinal, "0.00"))
            Tb("pRedBCST") = CDbl(IIf(vPRedBCST = "", 0, Format(vPRedBCST, "@")))
            Tb("vBCST") = CDbl(Format(curVBCST, "0.00"))
            Tb("pICMSST") = CDbl(Format(dblAliqInterna, "0.00"))
            Tb("vICMSST") = CDbl(Format(curVICMSST, "0.00"))
        End If
    Else
        ' chkICMSST desmarcado: zera todos os campos ST
        Tb("modBCST") = Format(0, "@")
        Tb("pMVAST") = CDbl(Format(0, "@"))
        Tb("pRedBCST") = CDbl(Format(0, "@"))
        Tb("vBCST") = CDbl(Format(0, "@"))
        Tb("pICMSST") = CDbl(Format(0, "@"))
        Tb("vICMSST") = CDbl(Format(0, "@"))
    End If
    

    
    
'    If txtValorProdICMS.Text = "" Then Tb("vICMS") = CDbl(Format(0, "@")) Else Tb("vICMS") = CDbl(Format(txtValorProdICMS, "@"))
    'If txtValorProdIPI.Text = "" Then Tb("IPIvIPI") = CDbl(Format(0, "@")) Else Tb("IPIvIPI") = CDbl(Format(txtValorProdIPI, "@"))
    
    'If txtICMS.Text <> "" Then Tb("pICMS") = CDbl(Format(txtICMS, "@"))
    'If txtICMS.Text <> "" Then Tb("vBC") = CDbl(Format(txtSubTotal, "@"))
    ' DIFAL: preenche valores corretos no INSERT
    If cboDestOperacao.Text = "2 - Operação Interestadual" Then
        If cboConsumidorFinal.Text = "1 - SIM" Then

            ' 1. Aliquota interestadual (origem x destino)
            sSQL = "SELECT AliquotaInterestadual FROM TribMatrizInterestadual WHERE UF_Origem = '" & vUFEmpresa & "' AND UF_Destino = '" & vUFDest & "'"
            Set rDifalLDI = dbData.OpenRecordset(sSQL)
            If Not rDifalLDI.EOF Then vPICMSInter = rDifalLDI("AliquotaInterestadual")

            ' 2. Regras do estado de destino (vigente)
            sSQL = "SELECT TOP 1 AliquotaInterna, AliquotaFCP, TipoCalculo, FCPCompoeBase FROM TribRegraDifalUF " & _
                   "WHERE UF_Destino = '" & vUFDest & "' AND DataInicioVigencia <= GETDATE() " & _
                   "AND (DataFimVigencia IS NULL OR DataFimVigencia >= GETDATE()) " & _
                   "ORDER BY DataInicioVigencia DESC"
            Set rDifalLDI = dbData.OpenRecordset(sSQL)
            If Not rDifalLDI.EOF Then
                vPICMSUFDest = rDifalLDI("AliquotaInterna")
                vPFCPUFDest = rDifalLDI("AliquotaFCP")
                vTipoCalcLDI = rDifalLDI("TipoCalculo")
                vFCPBaseLDI = rDifalLDI("FCPCompoeBase")
            End If

            ' 3. Base de calculo (base dupla ou simples)
            vBaseItem = CDbl(txtSubTotal.Text)
            ' Inclui IPI na base do DIFAL se parametro ativo (Art. 13 Lei Kandir - consumidor final)
            If vIPICompoeDIFAL = 1 Then vBaseItem = vBaseItem + CDbl(vValorIPI)
            If vTipoCalcLDI = 2 Then
                If vFCPBaseLDI = 1 Then
                    vBCUFDestLDI = (vBaseItem - (vBaseItem * vPICMSInter / 100)) / (1 - (vPICMSUFDest + vPFCPUFDest) / 100)
                Else
                    vBCUFDestLDI = (vBaseItem - (vBaseItem * vPICMSInter / 100)) / (1 - vPICMSUFDest / 100)
                End If
            Else
                vBCUFDestLDI = vBaseItem - (vBaseItem * vPICMSInter / 100)
            End If

            ' 4. DIFAL e FCP
            vICMSUFDestLDI = vBCUFDestLDI * (vPICMSUFDest - vPICMSInter) / 100
            vFCPUFDestLDI = vBCUFDestLDI * vPFCPUFDest / 100

            Tb("pICMSInter") = vPICMSInter
            Tb("pICMSUFDest") = vPICMSUFDest
            Tb("pFCPUFDest") = vPFCPUFDest
            Tb("pICMSInterPart") = 100
            Tb("vICMSUFRemet") = 0
            Tb("vBCUFDest") = vBCUFDestLDI
            Tb("vBCFCPUFDest") = vBCUFDestLDI
            Tb("vICMSUFDest") = vICMSUFDestLDI
            Tb("vFCPUFDest") = vFCPUFDestLDI

        Else
            Tb("vBCUFDest") = 0: Tb("vBCFCPUFDest") = 0: Tb("pFCPUFDest") = 0
            Tb("pICMSUFDest") = 0: Tb("pICMSInter") = 0: Tb("pICMSInterPart") = 0
            Tb("vFCPUFDest") = 0: Tb("vICMSUFRemet") = 0: Tb("vICMSUFDest") = 0
        End If
    Else
        Tb("vBCUFDest") = 0: Tb("vBCFCPUFDest") = 0: Tb("pFCPUFDest") = 0
        Tb("pICMSUFDest") = 0: Tb("pICMSInter") = 0: Tb("pICMSInterPart") = 0
        Tb("vFCPUFDest") = 0: Tb("vICMSUFRemet") = 0: Tb("vICMSUFDest") = 0
    End If
End Sub

Private Sub Calcular_Total()
Dim var_Quant As Double
Dim var_VALOR As Currency, var_Total As Currency

If txtQuant.Text = "" Then var_Quant = 1 Else var_Quant = txtQuant.Text
If txtValor.Text = "" Then var_VALOR = 0 Else var_VALOR = txtValor.Text

var_Total = var_VALOR * var_Quant
txtSubTotal.Text = FormatNumber(var_Total, 2)
End Sub
Sub Clear_Controls()
'Limpa_Tudo Me
mskEmissao.Mask = ""
mskSaida.Mask = ""
mskHora.Mask = ""
mskEmissao.Text = ""
mskSaida.Text = ""
mskHora.Text = ""
MostraStatus.Caption = ""
End Sub

Private Sub LimparObjetosDuplicata()
txtNumDup.Text = ""
txtTotalDup.Text = Format(0, ocMONEY)
txtNumParcDup.Text = "1"
txtIntervaloDup.Text = "30"
mskInicioDup.Mask = ""
mskInicioDup.Text = ""
txtValorParcDup.Text = Format(0, ocMONEY)
End Sub

Private Sub LimparObjetosProduto()
txtCodBarra.Text = ""
cboDescricao.Text = ""
txtCodProduto.Text = ""
txtValor = Format("0", "@")
txtSubTotal = Format("0", "@")
txtQuant = Format("1", "@")
txtDesc = Format("0", "@")
txtFrete = Format("0", "@")
txtSeguro = Format("0", "@")
txtOutrosItem = Format("0", "@")
End Sub


Sub Load_Data()
On Error GoTo erro
    vgDb.BeginTrans
    TbNotas("ChavedeAcessoAdicional") = Format(txtChaveReferenciada.Text, "@")
    If txtVolPesoBruto.Text = "" Then txtVolPesoBruto.Text = "0"
    If txtVolPesoLiquido.Text = "" Then txtVolPesoLiquido.Text = "0"
    TbNotas("CodigoNatureza") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    
    TbNotas("CodigoNota") = Format(txtCodNota.Text, "@")
    
    TbNotas("SerieNF") = Format(txtSerie.Text, "@")

    TbNotas("TipoCliente") = Format(cboTipoDest, "@")
    TbNotas("InformacoesComplementares") = txtInfComple
    'TbNotas("InformacoesComplementares") = Format(txtInfComple, "@")
    TbNotas("CFOP") = Format(cboNatureza.Text, "@")
    TbNotas("NaturezaOperacao") = Format(Left(txtNatureza, 59), "@")
    TbNotas("TipoDocumento") = IIf(IsNull(Format(Left(cboTipoNota.Text, 1), "@")) Or Vazio(Format(Left(cboTipoNota.Text, 1), "@")), 1, Format(Left(cboTipoNota.Text, 1), "@"))
    TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")

    TbNotas("DataEmissao") = IIf(txtCodPedido.Text <> "0", IIf(TbNotas("DataEmissao") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskEmissao, "@")), Format(mskEmissao, "@"))
    TbNotas("DataSaida") = IIf(txtCodPedido.Text <> "0", IIf(TbNotas("DataSaida") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskSaida, "@")), Format(mskSaida, "@"))
    TbNotas("HoraSaida") = IIf(txtCodPedido.Text <> "0", Format(Time(), "HH:MM:ss"), Format(mskHora, "@"))
    
    TbNotas("FinalidadeEmissaoNFe") = Format(cboFinalidade, "@")
    TbNotas("CodigoObservacao") = IIf(IsNull(Format(txtCodObservacao, "@")) Or Vazio(Format(txtCodObservacao, "@")), 0, Format(txtCodObservacao, "@"))
    TbNotas("NumeroNota") = Format(txtNumNota, "@")
    TbNotas("cCodigoNota") = IIf(TbNotas("cCodigoNota") = 0, GeraCodigoNota, TbNotas("cCodigoNota"))
    
    TbNotas("IndicadorFormaPagamento") = Format(cboIndicadorPagamento.Text, "@")
    TbNotas("FormaPagamento") = IIf(IsNull(Format(cboFormaPgto, "@")) Or Vazio(Format(cboFormaPgto, "@")), "01 = Dinheiro", Format(cboFormaPgto, "@"))
    'TbNotas("FormaPagamento") = Format(cboFormaPgto.Text, "@")
    
    TbNotas("CodigoCorrentista") = IIf(IsNull(Format(txtCodCliente, "@")) Or Vazio(Format(txtCodCliente, "@")), 0, Format(txtCodCliente, "@"))
    TbNotas("RazaoSocial") = Format(cboCliente, "@")
    
    'cboFormaPgto
    TbNotas("FormatoImpressaoDANFE") = Format(cboFormatoDANFe.Text, "@")
    
    TbNotas("FormatoEmissaoNFe") = Format(cboTipoEmissao.Text, "@")
    
    If Left(cboTipoEmissao.Text, 1) <> "1" Then
        TbNotas("ContingenciaDataHora") = Format(Now, "yyyy-mm-ddThh:mm:ss") & UTC
        TbNotas("ContingenciaJustificativa") = "EMISSAO DE NFE EM CONTIGENCIA DEVIDO A INDISPONIBILIDADE DO SERVICO NORMAL"
    End If
        
    TbNotas("IdentificadorDestino") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("IndicadorIEDestinatario") = IIf(IsNull(Format(Left(cboTipoContribuinte.Text, 1), "@")) Or Vazio(Format(Left(cboTipoContribuinte.Text, 1), "@")), 1, Format(Left(cboTipoContribuinte.Text, 1), "@"))
    TbNotas("ConsumidorFinal") = IIf(IsNull(Format(Left(cboConsumidorFinal.Text, 1), "@")) Or Vazio(Format(Left(cboConsumidorFinal.Text, 1), "@")), 1, Format(Left(cboConsumidorFinal.Text, 1), "@"))

    'tributos e valores
    TbNotas("BaseICMS") = IIf(IsNull(Format(txtBaseICMS, "@")) Or Vazio(Format(txtBaseICMS, "@")), 0, CDbl(Format(txtBaseICMS, "##0.00")))
    TbNotas("BaseICMSST") = IIf(IsNull(Format(txtBaseICMSST, "@")) Or Vazio(Format(txtBaseICMSST, "@")), 0, CDbl(Format(txtBaseICMSST, "##0.00")))
    TbNotas("ValorFrete") = IIf(Vazio(txtValorFrete), 0, CDbl(Format(txtValorFrete, "##0.00")))
    TbNotas("ValorSeguro") = IIf(IsNull(Format(txtValorSeguro, "@")) Or Vazio(Format(txtValorSeguro, "@")), 0, CDbl(Format(txtValorSeguro, "##0.00")))
    TbNotas("ValorOutrasDespesas") = IIf(IsNull(Format(txtValorOutrasDespesas, "@")) Or Vazio(Format(txtValorOutrasDespesas, "@")), 0, CDbl(Format(txtValorOutrasDespesas, "##0.00")))
    TbNotas("ValorICMS") = IIf(IsNull(Format(txtValorICMS, "@")) Or Vazio(Format(txtValorICMS, "@")), 0, CDbl(Format(txtValorICMS, "##0.000")))
    TbNotas("ValorICMSST") = IIf(IsNull(Format(txtValorICMSST, "@")) Or Vazio(Format(txtValorICMSST, "@")), 0, CDbl(Format(txtValorICMSST, "##0.00")))
    TbNotas("ValorIPI") = IIf(IsNull(Format(txtValorIPI, "@")) Or Vazio(Format(txtValorIPI, "@")), 0, CDbl(Format(txtValorIPI, "##0.000")))
    TbNotas("ValorProdutos") = IIf(IsNull(Format(txtTotaldosProdutos, "@")) Or Vazio(Format(txtTotaldosProdutos, "@")), 0, CDbl(FormatNumber(txtTotaldosProdutos, 2)))
    TbNotas("ValorDesconto") = IIf(IsNull(Format(txtValorDesconto, "@")) Or Vazio(Format(txtValorDesconto, "@")), 0, CDbl(Format(txtValorDesconto, "##0.00")))

    TbNotas("valornota") = IIf(IsNull(Format(txtTotaldaNota, "@")) Or Vazio(Format(txtTotaldaNota, "@")), 0, CDbl(FormatNumber(txtTotaldaNota, 2)))

    'TbNotas("BaseICMS") = " & FSQL(txtBaseICMS, 2) & "
    'TbNotas("BaseICMSST") = " & FSQL(txtBaseICMSST, 2) & "
    'TbNotas("ValorFrete") = Format(txtValorFrete.Text, "@")
    'TbNotas("ValorSeguro") = Format(txtValorSeguro.Text, "@")
    'TbNotas("ValorOutrasDespesas") = Format(txtValorOutrasDespesas.Text, "@")
    'TbNotas("ValorICMS") = Format(txtValorICMS.Text, "@")
    'TbNotas("ValorICMSST") = Format(txtValorICMSST.Text, "@")
    'TbNotas("ValorIPI") = Format(txtValorIPI.Text, "@")
    'TbNotas("ValorDesconto") = Format(txtValorDesconto.Text, "@")

    'transporte
    TbNotas("ModFrete") = IIf(IsNull(Format(Left(cboModFrete.Text, 1), "@")) Or Vazio(Format(Left(cboModFrete.Text, 1), "@")), 9, Format(Left(cboModFrete.Text, 1), "@"))
    TbNotas("TranspCodigo") = IIf(IsNull(Format(txtCodTransporte, "@")) Or Vazio(Format(txtCodTransporte, "@")), 0, Format(txtCodTransporte, "@"))
    TbNotas("TranspNome") = Format(cboTransporte, "@")
    TbNotas("TranspCNPJ_CPF") = Format(vTranspCNPJ, "@")
    TbNotas("TranspEndereco") = Format(vTranspEnd, "@")
    TbNotas("TranspMunicipio") = Format(vTranspCidade, "@")
    TbNotas("TranspUF") = Format(vTranspUF, "@")
    TbNotas("TranspInscricaoEstadual") = Format(vTranspIE, "@")
    
    TbNotas("TranspPlaca") = Format(txtPlaca, "@")
    TbNotas("TranspPlacaUF") = Format(txtPlacaUF, "@")
    
    TbNotas("VolumeQuantidade") = Format(txtVolQuant, "@")
    TbNotas("VolumeEspecie") = Format(txtVolEspecie, "@")
    TbNotas("VolumeMarca") = Format(txtVolMarca, "@")
    TbNotas("VolumeNumeracao") = Format(txtVolNumeracao, "@")
    TbNotas("VolumePesoBruto") = IIf(IsNull(Format(txtVolPesoBruto, "@")) Or Vazio(Format(txtVolPesoBruto, "@")), 0, CDbl(Format(txtVolPesoBruto, "##0.000")))
    TbNotas("VolumePesoLiquido") = IIf(IsNull(Format(txtVolPesoLiquido, "@")) Or Vazio(Format(txtVolPesoLiquido, "@")), 0, CDbl(Format(txtVolPesoLiquido, "##0.000")))

    'fatura
    TbNotas("NumeroFatura") = Format(txtNumNota, "@")
    TbNotas("ValorOriginalFatura") = IIf(IsNull(Format(txtTotaldosProdutos, "@")) Or Vazio(Format(txtTotaldosProdutos, "@")), 0, CDbl(FormatNumber(txtTotaldosProdutos, 2)))
    TbNotas("ValorDescontoFatura") = IIf(IsNull(Format(txtValorDesconto, "@")) Or Vazio(Format(txtValorDesconto, "@")), 0, CDbl(Format(txtValorDesconto, "##0.00")))
    TbNotas("ValorLiquidoFatura") = IIf(IsNull(Format(txtTotaldaNota, "@")) Or Vazio(Format(txtTotaldaNota, "@")), 0, CDbl(FormatNumber(txtTotaldaNota, 2)))

    TbNotas("InscricaoEstadual") = 0
    TbNotas("Suframa") = 0
    TbNotas("CNPJ_CPF") = 0
    TbNotas("Logradouro") = 0
    TbNotas("numero") = 0
    TbNotas("CodigoIBGE") = 0
    TbNotas("Bairro") = 0
    TbNotas("Complemento") = 0
    TbNotas("Municipio") = 0
    TbNotas("UF") = 0
    TbNotas("CEP") = 0
    TbNotas("CODIGOPAIS") = 0
    TbNotas("PAIS") = 0
    TbNotas("TELEFONE") = 0
    TbNotas("Inutilizada") = 0
    Exit Sub

Resume

erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
    vgDb.RollbackTrans
    Exit Sub
End Sub
Public Sub Load_Controls()
'On Error GoTo erro
    cboTipoDest = Format(TbNotas("TipoCliente"), "@")
    txtCodCliente = Format(TbNotas("CodigoCorrentista"), "@")
    cboCliente = Format(TbNotas("RazaoSocial"), "@")
    
    If TbNotas("CodigoNatureza") = 1 Then
        cboDestOperacao.Text = "1 - Operação Interna"
    ElseIf TbNotas("CodigoNatureza") = 2 Then
        cboDestOperacao.Text = "2 - Operação Interestadual"
    ElseIf TbNotas("CodigoNatureza") = 3 Then
        cboDestOperacao.Text = "3 - Operação com Exterior"
    End If

    If TbNotas("IndicadorIEDestinatario") = 1 Then
        cboTipoContribuinte.Text = "1 - CONTRIBUINTE ICMS"
    ElseIf TbNotas("IndicadorIEDestinatario") = 2 Then
        cboTipoContribuinte.Text = "2 - CONTRIBUINTE ISENTO"
    ElseIf TbNotas("IndicadorIEDestinatario") = 9 Then
        cboTipoContribuinte.Text = "9 - NÃO CONTRIBUINTE"
    End If
    
    If TbNotas("ConsumidorFinal") = False Then
        cboConsumidorFinal.Text = "0 - NÃO"
    ElseIf TbNotas("ConsumidorFinal") = True Then
        cboConsumidorFinal.Text = "1 - SIM"
    End If

'cboConsumidorFinal

    txtSerie = Format(TbNotas("SerieNF"), "@")
    cboFinalidade = Format(TbNotas("FinalidadeEmissaoNFe"), "@")
    txtInfAdicionais = Format(TbNotas("InformacoesAdicionais"), "@")
    cboNatureza = Format(TbNotas("CFOP"), "@")
    txtNatureza = Format(TbNotas("NaturezaOperacao"), "@")
    mskEmissao = Format(TbNotas("DataEmissao"), "@")
    'mskDataProtocolo = Format(TbNotas("DataHoraProcotolo"), "@")
    mskSaida = Format(TbNotas("DataSaida"), "@")
    mskHora = Format(TbNotas("HoraSaida"), "@")
    txtCodTransporte = Format(TbNotas("TranspCodigo"), "@")
    cboTransporte = Format(TbNotas("TranspNome"), "@")
    txtPlaca = Format(TbNotas("TranspPlaca"), "@")
    txtVolQuant = Format(TbNotas("VolumeQuantidade"), "@")
    txtVolEspecie = Format(TbNotas("VolumeEspecie"), "@")
    txtVolMarca = Format(TbNotas("VolumeMarca"), "@")
    txtVolNumeracao = Format(TbNotas("VolumeNumeracao"), "@")
    txtCodObservacao = Format(TbNotas("CodigoObservacao"), "@")
    txtNumNota = Format(TbNotas("NumeroNota"), "@")
    
    txtTotaldosProdutos = FormatNumber(TbNotas("ValorProdutos"), 2) '
    txtValorSeguro = FormatNumber(TbNotas("ValorSeguro"), 2)
    txtValorOutrasDespesas = FormatNumber(TbNotas("ValorOutrasDespesas"), 2)
    txtValorFrete = FormatNumber(TbNotas("ValorFrete"), 2)
    txtBaseICMS = FormatNumber(TbNotas("BaseICMS"), 2)
    txtValorICMS.Text = FormatNumber(TbNotas("ValorICMS"), 2)
    txtBaseICMSST = FormatNumber(TbNotas("BaseICMSST"), 2)
    txtValorICMSST.Text = FormatNumber(TbNotas("ValorICMSST"), 2)
    txtValorIPI.Text = FormatNumber(TbNotas("ValorIPI"), 2)
    txtValorDesconto.Text = FormatNumber(TbNotas("ValorDesconto"), 2)
    txtTotaldaNota = FormatNumber(TbNotas("ValorNota"), 2) '
    
    txtVolPesoBruto = Format(TbNotas("VolumePesoBruto"), "@")
    txtVolPesoLiquido = Format(TbNotas("VolumePesoLiquido"), "@")
    txtPlacaUF = Format(TbNotas("TranspPlacaUF"), "@")

    If TbNotas("ModFrete") = 0 Then
        cboModFrete.Text = "0 - Frete por conta do Remetente (CIF)"
    ElseIf TbNotas("ModFrete") = 1 Then
        cboModFrete.Text = "1 - Frete por conta do Destinatário (FOB)"
    ElseIf TbNotas("ModFrete") = 2 Then
        cboModFrete.Text = "2 - Frete por conta de Terceiros"
    ElseIf TbNotas("ModFrete") = 3 Then
        cboModFrete.Text = "3 - Transporte Próprio por conta do Remetente"
    ElseIf TbNotas("ModFrete") = 4 Then
        cboModFrete.Text = "4 - Transporte Próprio por conta do Destinatário"
    ElseIf TbNotas("ModFrete") = 9 Then
        cboModFrete.Text = "9 - Sem Ocorrência de Transporte"
    End If

    If TbNotas("TipoDocumento") = 0 Then
        cboTipoNota.Text = "0 - ENTRADA"
    ElseIf TbNotas("TipoDocumento") = 1 Then
        cboTipoNota.Text = "1 - SAÍDA"
    End If
    
    txtCodNota = Format(TbNotas("CodigoNota"), "@")
    'Text30 = Format(TbNotas("ChavedeAcesso"), "@")
    'Text31 = Format(TbNotas("NumeroProtocolo"), "@")
    'Text32 = Format(TbNotas("NumeroRecibo"), "@")
    cboIndicadorPagamento.Text = Format(TbNotas("IndicadorFormaPagamento"), "@") '"2 - Outros"
    cboFormaPgto.Text = Format(TbNotas("FormaPagamento"), "@")
    
    'TbNotas("IndicadorFormaPagamento") = Format(cboIndicadorPagamento.Text, "@")
    'TbNotas("IndicadorFormaPagamento") = Format(cboFormaPgto.Text, "@")
    
    cboFormatoDANFe.Text = Format(TbNotas("FormatoImpressaoDANFE"), "@")
    cboTipoEmissao.Text = Format(TbNotas("FormatoEmissaoNFe"), "@")
    txtCodPedido = Format(TbNotas("cod_pedido"), "@")
    'txtInfComple.Text = Format(TbNotas("InformacoesComplementares"), "@")
    txtInfComple.Text = TbNotas("InformacoesComplementares")
    txtChaveReferenciada.Text = Format(TbNotas("ChavedeAcessoAdicional"), "@")
    
    'cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    'cboFormaPgto.Text = "01 = Dinheiro"  'se não exisitir parcelas
    
    'fatura
    txtNumFatura.Text = Format(TbNotas("NumeroFatura"), "@")
    txtSubtotalFatura.Text = Format(TbNotas("ValorOriginalFatura"), ocMONEY)
    txtDescFatura.Text = Format(TbNotas("ValorDescontoFatura"), ocMONEY)
    txtTotalFatura.Text = Format(TbNotas("ValorLiquidoFatura"), ocMONEY) '

    'If vTipoEdicaoNFe = "Edicao" Then
        MostraStatus = MostraStatus_F9()
    '    frmItens.Enabled = True
    'End If
    'frmNota.Enabled = True
    'frmTransmissao.Enabled = True
    
    'If Text30.Text <> "" Then cmdConsultar.Enabled = True
    
    Mostrar_AliqUF
    AplicarEstadoCheckboxes
Exit Sub

Resume

'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub Mostrar_AliqUF()
sSQL = "SELECT CRT, ESTADO, RegimeTributario, IPICompoeDIFAL FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    vTipoCRT = r("CRT")
    vUFEmpresa = r("ESTADO")
    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))
    vIPICompoeDIFAL = IIf(IsNull(r("IPICompoeDIFAL")), 0, r("IPICompoeDIFAL"))
    
    If Left(cboDestOperacao.Text, 1) = 2 Then
        vAliqUFInter = Format(12, "#0.00")
        vAliqUFDest = Format(18, "#0.00")
    Else
        vAliqUFInter = Format(0, "#0.00")
        vAliqUFDest = Format(0, "#0.00")
    End If
End If
End Sub

Private Sub Mostrar_ItensNota()
Dim enviada As Boolean
Dim totalRegistros As Long
    
'On Error GoTo ErrLoad

sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, pICMS, vICMS, ValorUnitarioComercializacao, QuantidadeComercial, valordesconto, ValorTotalBruto, IPIpIPI, IPIvIPI FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

If Tb.RecordCount > 0 Then totalRegistros = Tb.RecordCount


'    enviada = SQLExecutaRetorno("SELECT Enviada FROM NotaFiscal WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text), "Enviada", 0)
'
'    If enviada Then
'       cboDestOperacao.Enabled = False
'       txtValorIPI .Enabled = False
'       mskEmissao.Enabled = False
'       mskSaida.Enabled = False
'       mskHora.Enabled = False
'       Text7.Enabled = False
'       Text8.Enabled = False
'       txtPlaca.Enabled = False
'       txtVolQuant.Enabled = False
'    End If

Exibir_Itens
Tab_Produtos.Enabled = True
Exit Sub
    
'ErrLoad:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
End Sub

Private Sub FormatarGridDuplicatas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With Grid_Duplicata
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 5
      .rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 500
      .ColWidth(2) = 2000
      .ColWidth(3) = 2000
      .ColWidth(4) = 2000
           

      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "FORMA"
      .TextMatrix(0, 3) = "VENC."
      .TextMatrix(0, 4) = "VALOR"

      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 7
      .ColAlignment(2) = 7
      .ColAlignment(3) = 7
      .ColAlignment(4) = 7
  
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("Sequencia"), "000")
            .TextMatrix(.rows - 1, 2) = rTabela("var_FormaPgto")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("Vencimento"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 4) = Format(rTabela("ValorDocumento"), ocMONEY)

            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      .rows = .rows - 1
      
      .Visible = True
      .Redraw = True
   End With
End Sub
Private Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim j As Integer

   With GridNotasItens
      .Visible = False
      .Redraw = False

      .Clear
      .Cols = 27
      .rows = 2

      'Colunas fixas (sempre visiveis)
      .ColWidth(0) = 200    'indicador de linha
      .ColWidth(1) = 400    'No.
      .ColWidth(2) = 1500   'EAN
      .ColWidth(3) = 0      'COD. (oculto)
      .ColWidth(4) = 3500   'DESCRICAO
      .ColWidth(5) = 450    'UND
      .ColWidth(6) = 900    'NCM
      .ColWidth(7) = 600    'CFOP
      .ColWidth(8) = 500    'CST
      .ColWidth(9) = 850    'VALOR
      .ColWidth(10) = 850   'QTDE
      .ColWidth(11) = 800   'FRETE
      .ColWidth(12) = 900   'SEGURO
      .ColWidth(13) = 900   'OUTROS
      .ColWidth(14) = 800   'DESC.
      .ColWidth(15) = 1050  'TOTAL
      'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)
      .ColWidth(16) = 0     'BC ICMS
      .ColWidth(17) = 0     '%ICMS
      .ColWidth(18) = 0     'ICMS
      .ColWidth(19) = 0     '%RED BC
      .ColWidth(20) = 0     'BC ST
      .ColWidth(21) = 0     '%ICMSST
      .ColWidth(22) = 0     'ICMSST
      .ColWidth(23) = 0     'MVA ST
      .ColWidth(24) = 0     '%IPI
      .ColWidth(25) = 0     'IPI
      .ColWidth(26) = 0     'cEnq

      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "EAN"
      .TextMatrix(0, 3) = "CÓD."
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "UND"
      .TextMatrix(0, 6) = "NCM"
      .TextMatrix(0, 7) = "CFOP"
      .TextMatrix(0, 8) = "CST"
      .TextMatrix(0, 9) = "VALOR"
      .TextMatrix(0, 10) = "QTDE"
      .TextMatrix(0, 11) = "FRETE"
      .TextMatrix(0, 12) = "SEGURO"
      .TextMatrix(0, 13) = "OUTROS"
      .TextMatrix(0, 14) = "DESC."
      .TextMatrix(0, 15) = "TOTAL"
      .TextMatrix(0, 16) = "BC ICMS"
      .TextMatrix(0, 17) = "%ICMS"
      .TextMatrix(0, 18) = "ICMS"
      .TextMatrix(0, 19) = "%RED BC"
      .TextMatrix(0, 20) = "BC ST"
      .TextMatrix(0, 21) = "%ICMSST"
      .TextMatrix(0, 22) = "ICMSST"
      .TextMatrix(0, 23) = "MVA ST"
      .TextMatrix(0, 24) = "CST IPI"
      .TextMatrix(0, 25) = "%IPI"
      .TextMatrix(0, 26) = "IPI"

      'Cabecalho em negrito e centralizado
      For i = 0 To .Cols - 1
         .Col = i: .Row = 0
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next i

      'Alinhamento: texto esquerda (0-8), numeros direita (9-26)
      For i = 0 To 8
         .ColAlignment(i) = 1
      Next i
      For i = 9 To 26
         .ColAlignment(i) = 6
      Next i

      i = 1
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("ITEM"), "000")
            .TextMatrix(.rows - 1, 2) = rTabela("EAN")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("CodigoProduto"), "00000")
            .TextMatrix(.rows - 1, 4) = rTabela("NomeProduto")
            .TextMatrix(.rows - 1, 5) = rTabela("UnidadeComercial")
            .TextMatrix(.rows - 1, 6) = rTabela("NCM")
            .TextMatrix(.rows - 1, 7) = rTabela("CFOP")
            .TextMatrix(.rows - 1, 8) = rTabela("CST")
            .TextMatrix(.rows - 1, 9) = FormatNumber(rTabela("ValorUnitarioComercializacao"), 2)
            If rTabela("UnidadeComercial") = "KG" Or rTabela("UnidadeComercial") = "GR" Or rTabela("UnidadeComercial") = "MG" Then
                .TextMatrix(.rows - 1, 10) = Format(rTabela("QuantidadeComercial"), ocPESO)
            Else
                .TextMatrix(.rows - 1, 10) = Format(rTabela("QuantidadeComercial"), "###,###,##0")
            End If
            .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("ValorFrete"), 2)
            .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("ValorSeguro"), 2)
            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela("ValorOutros"), 2)
            .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("ValorDesconto"), 2)
            .TextMatrix(.rows - 1, 15) = FormatNumber(rTabela("ValorTotalBruto"), 2)
            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela("vBC"), 2)
            .TextMatrix(.rows - 1, 17) = FormatNumber(rTabela("pICMS"), 2)
            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela("vICMS"), 2)
            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela("pRedBC"), 2)
            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela("vBCST"), 2)
            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela("pICMSST"), 2)
            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela("vICMSST"), 2)
            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela("pMVAST"), 2)
            .TextMatrix(.rows - 1, 24) = rTabela("IPICST")
            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("IPIpIPI"), 2)
            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela("IPIvIPI"), 2)

            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If

      .rows = .rows - 1

      'Numero da linha no col 0
      For i = 1 To .rows - 1
         .TextMatrix(i, 0) = i
      Next i

      'EAN em negrito
      For i = 1 To .rows - 1
         .Row = i: .Col = 2: .CellFontBold = True
      Next i

      'COD. em destaque
      For i = 1 To .rows - 1
         .Row = i: .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next i

      'TOTAL em destaque
      For i = 1 To .rows - 1
         .Row = i: .Col = 15
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next i

      'Colunas editáveis em amarelo claro
      Dim colEdit As Variant
      For Each colEdit In Array(2, 5, 6, 7, 8, 17, 19, 21, 23, 24, 25)
         For i = 1 To .rows - 1
            .Row = i: .Col = colEdit
            .CellBackColor = &HC8FFFF
         Next i
      Next colEdit

      GridNotasItens.Col = 0
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub AplicarEstadoCheckboxes()
    Dim bHabilitar As Boolean
    ' CRT 1/2/4 (Simples/MEI) com Finalidade <> 4: desabilita IPI, ST e RedBC
    bHabilitar = Not (vTipoCRT = 1 Or vTipoCRT = 2 Or vTipoCRT = 4) Or (Left(cboFinalidade.Text, 1) = "4")
    chkIPI.Enabled    = bHabilitar
    chkICMSST.Enabled = bHabilitar
    chkpRedBC.Enabled = bHabilitar
    If Not bHabilitar Then
        bSupressChkEvents = True
        chkIPI.Value    = 0
        chkICMSST.Value = 0
        chkpRedBC.Value = 0
        bSupressChkEvents = False
        AplicarVisibilidadeGridItens
    End If
End Sub

Sub AplicarVisibilidadeGridItens()
   If GridNotasItens.Cols < 27 Then Exit Sub
   'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)
   Dim bICMS As Boolean
   bICMS = (Left(cboFinalidade.Text, 1) = "4")
   GridNotasItens.ColWidth(16) = IIf(bICMS, 850, 0)  'BC ICMS
   GridNotasItens.ColWidth(17) = IIf(bICMS, 850, 0)  '%ICMS
   GridNotasItens.ColWidth(18) = IIf(bICMS, 850, 0)  'ICMS

   '%RedBC: chkpRedBC
   GridNotasItens.ColWidth(19) = IIf(chkpRedBC.Value = 1, 700, 0)

   'Grupo ICMSST: chkICMSST
   Dim bST As Boolean
   bST = (chkICMSST.Value = 1)
   GridNotasItens.ColWidth(20) = IIf(bST, 850, 0)  'BC ST
   GridNotasItens.ColWidth(21) = IIf(bST, 900, 0)  '%ICMSST
   GridNotasItens.ColWidth(22) = IIf(bST, 850, 0)  'ICMSST
   GridNotasItens.ColWidth(23) = IIf(bST, 850, 0)  'MVA ST

   'Grupo IPI: chkIPI
   Dim bIPI As Boolean
   bIPI = (chkIPI.Value = 1)
   GridNotasItens.ColWidth(24) = IIf(bIPI, 850, 0)  '%IPI
   GridNotasItens.ColWidth(25) = IIf(bIPI, 850, 0)  'IPI
   GridNotasItens.ColWidth(26) = IIf(bIPI, 850, 0)  'cEnq
End Sub



Private Sub chkIPI_Click()
    If bSupressChkEvents Then Exit Sub
    AplicarVisibilidadeGridItens
    If chkIPI.Value = 0 And txtCodNota.Text <> "" Then
        dbData.Execute "UPDATE NotaFiscalItens SET IPIcEnq = '999', IPIvBC = 0, IPIpIPI = 0, IPIvIPI = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
    End If
    RecalcularItensNota
End Sub

Private Sub chkpRedBC_Click()
    If bSupressChkEvents Then Exit Sub
    AplicarVisibilidadeGridItens
    If chkpRedBC.Value = 0 And txtCodNota.Text <> "" Then
        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)
    End If
    RecalcularItensNota
End Sub

Private Sub chkICMSST_Click()
    If bSupressChkEvents Then Exit Sub
    AplicarVisibilidadeGridItens
    RecalcularItensNota
End Sub

Private Sub cboFinalidade_Change()
    AplicarVisibilidadeGridItens
End Sub

Private Sub cboFinalidade_Click()
    AplicarVisibilidadeGridItens
    AplicarEstadoCheckboxes
    RecalcularItensNota
    CalcularICMSInterItensGERAL
End Sub

Private Sub cboFinalidade_LostFocus()
   AplicarVisibilidadeGridItens
End Sub

Private Sub cboConsumidorFinal_Click()
    RecalcularItensNota
    CalcularICMSInterItensGERAL
End Sub


Private Sub PreencherGridNotas()
Dim totalRegistros As Long

On Error GoTo ErrLoad

RsOpen TbConsulta, "SELECT *,  " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM NotaFiscal order by NumeroNota desc"
                
If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub SomarProdutosNota()
'Dim sSQL As String, vTotal As Double
'Dim var_ValorFrete
'var_ValorFrete = txtValorFrete.Text

'On Error GoTo erro

'    sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
'    vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
'    sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " + " & FSQL(var_ValorFrete, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
'    SQLExecuta sSQL
    
'    Exit Sub
'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub TransformarPedidoemNFE()
Dim tblItensPedido As ADODB.Recordset

'Atualiza a base de dados (funcionando)
Dim VarCodNota As Integer
VarCodNota = CInt(txtCodNota.Text)

'verificar se o pedido possui item com zero
sSQL = "SELECT COD_PRODUTO,  item FROM  pedidos_itens WHERE COD_PEDIDO = " & txtCodPedido.Text & ";"
Set r = dbData.OpenRecordset(sSQL)

Dim vItem As Integer
vItem = 1

If r!item = "0" Then
    For i = 1 To r.RecordCount
        dbData.Execute "UPDATE pedidos_itens SET item = " & vItem & " WHERE COD_PRODUTO = " & r!COD_PRODUTO & " and COD_PEDIDO = " & txtCodPedido.Text & ";"
    vItem = vItem + 1
    r.MoveNext
    Next
End If

sSQL = "INSERT INTO NotaFiscalItens ( " & _
        "CodigoProduto, " & _
        "EAN, " & _
        "NomeProduto, " & _
        "CFOP, " & _
        "NCM, " & _
        "CST, " & _
        "UnidadeComercial, " & _
        "ValorUnitarioComercializacao, " & _
        "ValorTotalBruto, " & _
        "tipodesconto, " & _
        "desconto, " & _
        "Valordesconto, " & _
        "QuantidadeComercial, " & _
        "pICMS, " & _
        "vBC, " & _
        "vICMS,  " & _
        "item, " & _
        "CodigoNota, TipoProduto " & _
        " ) " & _
        "SELECT pedidos_itens.cod_produto, produtos.EAN, produtos.descricao, produtos.cfop, produtos.ncm, produtos.icmscst, produtos.unid_medida, pedidos_itens.preco, pedidos_itens.Subtotal, 1, pedidos_itens.desconto, pedidos_itens.desconto, pedidos_itens.quantidade, 0, (pedidos_itens.Subtotal) as varVBC, 0, pedidos_itens.item, " & VarCodNota & ", (CASE produtos.combustivel WHEN 1 THEN 'Combustível' ELSE '' END) " & _
        "FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
        "WHERE pedidos_itens.COD_PEDIDO = " & txtCodPedido.Text & ";"
'Debug.Print sSQL
'sSQL = "INSERT INTO NotaFiscalItens ( " & _
        "CodigoProduto, " & _
        "EAN, " & _
        "NomeProduto, " & _
        "CFOP, " & _
        "NCM, " & _
        "CST, " & _
        "UnidadeComercial, " & _
        "ValorUnitarioComercializacao, " & _
        "ValorTotalBruto, " & _
        "tipodesconto, " & _
        "desconto, " & _
        "Valordesconto, " & _
        "QuantidadeComercial, " & _
        "pICMS, " & _
        "vBC, " & _
        "vICMS,  " & _
        "item, " & _
        "CodigoNota, TipoProduto " & _
        " ) " & _
        "SELECT pedidos_itens.cod_produto, produtos.EAN, produtos.descricao, produtos.cfop, produtos.ncm, produtos.icmscst, produtos.unid_medida, pedidos_itens.preco, (pedidos_itens.Subtotal) as varValorBruto, 1, pedidos_itens.desconto, pedidos_itens.desconto, pedidos_itens.quantidade, 0, (pedidos_itens.preco * pedidos_itens.quantidade) as varVBC, 0, pedidos_itens.item, " & VarCodNota & ", (CASE produtos.combustivel WHEN 1 THEN 'Combustível' ELSE '' END) " & _
        "FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
        "WHERE pedidos_itens.COD_PEDIDO = " & txtCodPedido.Text & ";"
dbData.Execute sSQL


'preencher o grid dos itens com o pedido
Exibir_Itens

'MOSTRAR A QUANTIDADE REGISTROS
'lblQuantPedidos.Caption = Format(totalRegistros, "00")
End Sub

Private Sub Verificar_Duplicatas()
sSQL = "SELECT * " & _
"FROM NotaFiscalParcelas " & _
"WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _
"ORDER BY Sequencia;"
RsOpen Tb, sSQL

If Not Tb.EOF Then
    Tb.MoveFirst
    If Tb("CodigoFormaPagamento") = 1 Then
        cboFormaPgto.Text = "01 = Dinheiro"
    ElseIf Tb("CodigoFormaPagamento") = 2 Then
        cboFormaPgto.Text = "02 = Cheque"
    ElseIf Tb("CodigoFormaPagamento") = 3 Then
        cboFormaPgto.Text = "03 = Cartão de Crédito"
    ElseIf Tb("CodigoFormaPagamento") = 4 Then
        cboFormaPgto.Text = "04 = Cartão de Débito"
    ElseIf Tb("CodigoFormaPagamento") = 5 Then
        cboFormaPgto.Text = "05 = Crédito Loja"
    ElseIf Tb("CodigoFormaPagamento") = 10 Then
        cboFormaPgto.Text = "10 = Vale Alimentação"
    ElseIf Tb("CodigoFormaPagamento") = 11 Then
        cboFormaPgto.Text = "11 = Vale Refeição"
    ElseIf Tb("CodigoFormaPagamento") = 12 Then
        cboFormaPgto.Text = "12 = Vale Presente"
    ElseIf Tb("CodigoFormaPagamento") = 13 Then
        cboFormaPgto.Text = "13 = Vale Combustível"
    ElseIf Tb("CodigoFormaPagamento") = 14 Then
        cboFormaPgto.Text = "14 = Duplicata Mercantil"
    ElseIf Tb("CodigoFormaPagamento") = 15 Then
        cboFormaPgto.Text = "15 = Boleto Bancário"
    ElseIf Tb("CodigoFormaPagamento") = 16 Then
        cboFormaPgto.Text = "16 = Depósito Bancário"
    ElseIf Tb("CodigoFormaPagamento") = 17 Then
        cboFormaPgto.Text = "17 = PIX"
    ElseIf Tb("CodigoFormaPagamento") = 18 Then
        cboFormaPgto.Text = "18 = Transferência bancária"
    ElseIf Tb("CodigoFormaPagamento") = 19 Then
        cboFormaPgto.Text = "19 = Programa de fidelidade"
    ElseIf Tb("CodigoFormaPagamento") = 29 Then
        cboFormaPgto.Text = "90 = Sem pagamento"
    ElseIf Tb("CodigoFormaPagamento") = 99 Then
        cboFormaPgto.Text = "99 = Outros"
    Else
        cboFormaPgto.Text = "01 = Dinheiro"
    End If
End If

If cboIndicadorPagamento.Text = "0 - Pagamento à vista" Then
    frmDuplicata.Visible = False
ElseIf cboIndicadorPagamento.Text = "1 - Pagamento à prazo" Then
    frmDuplicata.Visible = True
End If
End Sub

Private Sub VerificarDestinatarioEnviar()
vPossuiErro = False

Dim vTipoCliente As String
vCodCliente = (GridNotas.TextMatrix(GridNotas.Row, 12))
vTipoCliente = (GridNotas.TextMatrix(GridNotas.Row, 13))

If vTipoCliente = "FORNECEDOR" Then
    sSQL = "SELECT *, 'JURÍDICA' as vTipo  FROM FORNECEDOR WHERE codigo = " & Val(vCodCliente)
Else
    sSQL = "SELECT *, tipo as vTipo FROM cliente WHERE codigo = " & Val(vCodCliente)
End If

Set r = dbData.OpenRecordset(sSQL)

Dim vCPF As String
If Not r.EOF Then
vCPF = RemoverFormato(r("cpf"))
End If

'If ShowMsg("Deseja realmente transformar o pedido: " & GridPedidos.TextMatrix(GridPedidos.Row, 1) & " em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

If Not r.EOF And Not r.BOF Then
    If IsEmpty(r("endereco")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Endereço]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If IsEmpty(r("numero")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Número]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If IsEmpty(r("bairro")) Or Len(r("bairro")) < 4 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Bairro]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If IsEmpty(r("cidade")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Cidade]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If IsEmpty(r("estado")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Estado]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If IsEmpty(r("CodigoIBGE")) Or r("CodigoIBGE") = "0" Or Len(r("CodigoIBGE")) <> 7 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Cód IBGE]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If IsEmpty(r("CEP")) Or Len(r("CEP")) < 10 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: CEP]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    If r("TipoContribuinte") = 9 Then
        If IsEmpty(vCPF) Or Len(vCPF) < 11 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    Else
        If r("vTipo") = "RURAL" Then
            If IsEmpty(vCPF) Or Len(vCPF) < 11 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
        Else
            If IsEmpty(vCPF) Or Len(vCPF) < 14 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: CNPJ]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
        End If
    End If
    
    If r("TipoContribuinte") = 1 Then
        If Vazio(r("ie")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros no [Campo: Insc. Estadual]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente
    End If
End If

AtualizarCliente:
If vPossuiErro = True Then
    If cboTipoDest.Text = "FORNECEDOR" Then
        Load Fornecedor_Cadastro
        Fornecedor_Cadastro.SSTab1.Tab = 0
        Fornecedor_Cadastro.cmdNovo.Enabled = False
        Fornecedor_Cadastro.cmdSalvar.Enabled = False
        Fornecedor_Cadastro.cmdCancelar.Enabled = False
        Fornecedor_Cadastro.txtCodigo.Text = vCodCliente
        Fornecedor_Cadastro.Show 1
    Else
        Load Clientes_Cadastro
        Clientes_Cadastro.SSTab1.Tab = 0
        Clientes_Cadastro.cmdNovo.Enabled = False
        Clientes_Cadastro.cmdSalvar.Enabled = False
        Clientes_Cadastro.cmdCancelar.Enabled = False
        Clientes_Cadastro.cboTipoCliente.Text = "CADASTRO"
        Clientes_Cadastro.txtCodigo.Text = vCodCliente
        Clientes_Cadastro.Show 1
    End If
End If
End Sub
Private Sub VerificarDestinatario()
'Dim varCodCliente As String
vCodCliente = (GridNotas.TextMatrix(GridNotas.Row, 12))

If cboTipoDest.Text = "FORNECEDOR" Then
    sSQL = "SELECT * FROM FORNECEDOR WHERE codigo = " & Val(vCodCliente)
Else
    sSQL = "SELECT * FROM cliente WHERE codigo = " & Val(vCodCliente)
End If
Set r = dbData.OpenRecordset(sSQL)

Dim vCPF As String
If Not r.EOF Then
vCPF = RemoverFormato(r("cpf"))
End If

'If ShowMsg("Deseja realmente transformar o pedido: " & GridPedidos.TextMatrix(GridPedidos.Row, 1) & " em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

If Not r.EOF And Not r.BOF Then
    If IsEmpty(r("endereco")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Endereço]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If IsEmpty(r("numero")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Número]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If IsEmpty(r("bairro")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Bairro]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If IsEmpty(r("cidade")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Cidade]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If IsEmpty(r("estado")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Estado]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If IsEmpty(r("CodigoIBGE")) Or Len(r("CodigoIBGE")) < 7 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Cód IBGE]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If IsEmpty(r("CEP")) Or Len(r("CEP")) < 10 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: CEP]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    If r("TipoContribuinte") = 9 Then
        If IsEmpty(vCPF) Or Len(vCPF) < 11 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    Else
        If IsEmpty(vCPF) Or Len(vCPF) < 14 Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: CNPJ]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    End If
    
    If r("TipoContribuinte") = 1 Then
        If Vazio(r("ie")) Then If ShowMsg("O cadastro do DESTINATÁRIO possui erros [Campo: Insc. Estadual]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente
    End If
End If

Exit Sub

AtualizarCliente:
'varCodCliente = txtCodCliente.Text

If cboTipoDest.Text = "FORNECEDOR" Then
    Load Fornecedor_Cadastro
    Fornecedor_Cadastro.SSTab1.Tab = 0
    Fornecedor_Cadastro.cmdNovo.Enabled = False
    Fornecedor_Cadastro.cmdSalvar.Enabled = False
    Fornecedor_Cadastro.cmdCancelar.Enabled = False
    Fornecedor_Cadastro.txtCodigo.Text = vCodCliente
    Fornecedor_Cadastro.Show 1
Else
    Load Clientes_Cadastro
    Clientes_Cadastro.SSTab1.Tab = 0
    Clientes_Cadastro.cmdNovo.Enabled = False
    Clientes_Cadastro.cmdSalvar.Enabled = False
    Clientes_Cadastro.cmdCancelar.Enabled = False
    Clientes_Cadastro.cboTipoCliente.Text = "CADASTRO"
    Clientes_Cadastro.txtCodigo.Text = vCodCliente
    Clientes_Cadastro.Show 1
End If

End Sub


Private Sub VerificarProdutosEnviar()
'If GridNotas.rows <= 1 Then
'    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
'    Exit Sub
'End If

IdNFProd = GridNotas.TextMatrix(GridNotas.Row, 1)

'verificando os itens do pedido
sSQL = "SELECT CodigoNota, Item, CodigoProduto, NomeProduto, EAN, NCM, CFOP, CST, pICMS, vICMS, PISCST, PISpPIS, PISvPIS, COFINSCST, COFINSpCOFINS, COFINSvCOFINS, UnidadeComercial, QuantidadeComercial, ValorUnitarioComercializacao, ValorTotalBruto, vBC " & _
       "FROM NotaFiscalItens " & _
       "WHERE CodigoNota = " & IdNFProd
Set rNFCeItens = dbData.OpenRecordset(sSQL)

'Dim EncontroErroNFCe As Boolean
EncontroErroNFCe = False

 For i = 1 To rNFCeItens.RecordCount
     
     'NCM..........
     If rNFCeItens!EAN <> "SEM GTIN" Then
         If Len(rNFCeItens!EAN) > 13 Or Len(rNFCeItens!EAN) < 8 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com EAN incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If
    
     'CFOP..........
     If rNFCeItens!CFOP <> Empty Or rNFCeItens!CFOP = "0" Then
         If Len(rNFCeItens!CFOP) > 4 Or Len(rNFCeItens!CFOP) < 4 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com CFOP incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If
     
     'ICMS CST..........
     If rNFCeItens!CST <> Empty Then
         If Len(rNFCeItens!CST) > 3 Or Len(rNFCeItens!CST) < 3 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com ICMS CST incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If

     'PIS CST..........
     If rNFCeItens!PISCST <> Empty Then
         If Len(rNFCeItens!PISCST) > 2 Or Len(rNFCeItens!PISCST) < 2 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com PIS CST incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If

     'COFINS CST..........
     If rNFCeItens!COFINSCST <> Empty Then
         If Len(rNFCeItens!COFINSCST) > 2 Or Len(rNFCeItens!COFINSCST) < 2 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com COFINS CST incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If
     
     'NCM..........
     If rNFCeItens!NCM <> Empty Or rNFCeItens!NCM = "0" Or rNFCeItens!NCM = "" Then
         If Len(rNFCeItens!NCM) > 8 Or Len(rNFCeItens!NCM) < 8 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com NCM incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If
     'End If
     
     'UNIDADE DE MEDIDA..........
     If rNFCeItens!UnidadeComercial <> Empty Then
         If Len(rNFCeItens!UnidadeComercial) > 2 Or Len(rNFCeItens!UnidadeComercial) < 1 Then
             EncontroErroNFCe = True
         Else
             EncontroErroNFCe = False
         End If
     Else
         EncontroErroNFCe = False
     End If
     
     If EncontroErroNFCe = True Then 'GoTo Continuar
        If MsgBox("Produto com UNIDADE DE MEDIDA incorreto ou inválido! " & Chr(13) & " Produto: '" & rNFCeItens!NomeProduto & "' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then
            vPossuiErro = True
            GoTo Continuar
        Else
            vPossuiErro = True
            Exit Sub
        End If
    End If
 
 rNFCeItens.MoveNext
 Next
    
Continuar:
If EncontroErroNFCe = True Then
    vTipoEdicaoNFe = "Edicao"
    RsOpen TbNotas, "SELECT *,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE (CASE WHEN Inutilizada = 1 THEN 'Inutilizada' ELSE 'Em Digitação' END) END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & IdNFProd
    Load_Controls
    Frm_NF.Tab = 0
    cmdNovo.Enabled = False
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    frmNota.Enabled = True
    frmDestinatario.Enabled = True
    frmItens.Enabled = True
    Tab_Totais.Enabled = True
    Tab_Produtos.Enabled = True
    Exibir_Duplicatas
Else
    vPossuiErro = False
End If
End Sub


Private Sub cboAnoPedidos_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAnoPedidos.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAnoPedidos.AddItem i
Next

moCombo.AttachTo cboAnoPedidos
End Sub


Private Sub cboClientePedidos_GotFocus()
Dim r As ADODB.Recordset
'Dim itemAtual As String
'Dim codAtual As String

'itemAtual = CboCliente.Text
'codAtual = TxtCodCliente.Text
cboClientePedidos.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboClientePedidos.AddItem r("nome")
   cboClientePedidos.ItemData(cboClientePedidos.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'CboCliente.Text = itemAtual
'TxtCodCliente.Text = codAtual
moCombo.AttachTo cboClientePedidos
End Sub


Private Sub cboClientePedidos_LostFocus()
On Error GoTo TrataErro

If cboClientePedidos.Text = "" Then txtCodClientePedidos.Text = "": Exit Sub
If cboClientePedidos.ListIndex = -1 Then txtCodClientePedidos.Text = "": Exit Sub

txtCodClientePedidos = cboClientePedidos.ItemData(cboClientePedidos.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboConNotaAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboConNotaAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboConNotaAno.AddItem i
Next

moCombo.AttachTo cboConNotaAno
End Sub


Private Sub cboConNotaCliente_GotFocus()
If cboFiltroNota.Text = "NUM. NOTA" Then
    cboConNotaCliente.Text = ""
    cboConNotaCliente.Clear
Else
    'Dim itemAtual As String
    'Dim codAtual As String
    
    'itemAtual = CboCliente.Text
    'codAtual = TxtCodCliente.Text
    cboConNotaCliente.Clear
    
    sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboConNotaCliente.AddItem r("nome")
       cboConNotaCliente.ItemData(cboConNotaCliente.NewIndex) = r("codigo")
       r.MoveNext
    Loop
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    'CboCliente.Text = itemAtual
    'TxtCodCliente.Text = codAtual
    moCombo.AttachTo cboConNotaCliente
End If
End Sub


Private Sub cboConNotaCliente_LostFocus()
On Error GoTo TrataErro

If cboConNotaCliente.Text = "" Then txtConNotaCodCliente.Text = "": Exit Sub
If cboConNotaCliente.ListIndex = -1 Then txtConNotaCodCliente.Text = "": Exit Sub

txtConNotaCodCliente = cboConNotaCliente.ItemData(cboConNotaCliente.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboConNotaMes_GotFocus()
cboConNotaMes.Clear

cboConNotaMes.AddItem "Janeiro"
cboConNotaMes.AddItem "Fevereiro"
cboConNotaMes.AddItem "Março"
cboConNotaMes.AddItem "Abril"
cboConNotaMes.AddItem "Maio"
cboConNotaMes.AddItem "Junho"
cboConNotaMes.AddItem "Julho"
cboConNotaMes.AddItem "Agosto"
cboConNotaMes.AddItem "Setembro"
cboConNotaMes.AddItem "Outubro"
cboConNotaMes.AddItem "Novembro"
cboConNotaMes.AddItem "Dezembro"

moCombo.AttachTo cboConNotaMes
End Sub


Private Sub cboConsumidorFinal_GotFocus()
Dim VarText As String
VarText = cboConsumidorFinal.Text

cboConsumidorFinal.Clear
cboConsumidorFinal.AddItem "0 - NÃO"
cboConsumidorFinal.AddItem "1 - SIM"

If cboConsumidorFinal.Text = "" Then cboConsumidorFinal.Text = VarText
SelectControl cboConsumidorFinal
moCombo.AttachTo cboConsumidorFinal
End Sub


Private Sub cboDescricao_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim vUltimoValorVenda As String     '===================TER QUE COLOCAR DEPOIS PARA TODOS OS TIPOS DE VENDAS
vUltimoValorVenda = " (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) "
         
If Len(cboDescricao.Text) > 3 Then
    sSQL = "SELECT DISTINCT descricao, codigo FROM produtos WHERE (descricao LIKE '%" & cboDescricao.Text & "%') AND (produtos.ativo = 1) and " & vUltimoValorVenda & " > 0  ORDER BY descricao;"
    Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
       cboDescricao.AddItem ValidateNull(r("descricao"))
        'If cboDescricao.ListIndex <> -1 Then
            cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
        'End If
       r.MoveNext
    Loop
End If
End Sub

Private Sub AtualizarCFOPCSTItens()
    If txtCodNota.Text = "" Then Exit Sub
    Dim vCodNota As Long
    vCodNota = Val(txtCodNota.Text)
    Dim bInter As Boolean
    bInter = (Left(cboDestOperacao.Text, 1) = "2")

    ' 1. Converter CFOP
    If bInter Then
        dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '6' + SUBSTRING(CFOP, 2, 3) WHERE CodigoNota = " & vCodNota & " AND LEFT(CFOP, 1) = '5'"
    Else
        dbData.Execute "UPDATE NotaFiscalItens SET CFOP = '5' + SUBSTRING(CFOP, 2, 3) WHERE CodigoNota = " & vCodNota & " AND LEFT(CFOP, 1) = '6'"
    End If

    ' 2. Atualizar CST/CSOSN para Simples (regime 1 ou 2)
    If vRegimeTributario = 1 Or vRegimeTributario = 2 Then
        dbData.Execute "UPDATE NotaFiscalItens SET CST = CASE WHEN RIGHT(CFOP, 3) = '102' THEN '102' WHEN RIGHT(CFOP, 3) = '405' THEN '500' ELSE CST END WHERE CodigoNota = " & vCodNota
    End If

    ' 3. Recalcular impostos e exibir grid
    RecalcularItensNota
    Exibir_Itens
End Sub

Private Sub cboDestOperacao_Change()
cboDestOperacao_LostFocus
End Sub

Private Sub cboDestOperacao_GotFocus()
Dim VarText As String
VarText = cboDestOperacao.Text

cboDestOperacao.Clear
cboDestOperacao.AddItem "1 - Operação Interna"
cboDestOperacao.AddItem "2 - Operação Interestadual"
cboDestOperacao.AddItem "3 - Operação com Exterior"

If cboDestOperacao.Text = "" Then cboDestOperacao.Text = VarText
SelectControl cboDestOperacao
moCombo.AttachTo cboDestOperacao
End Sub


Private Sub cboDestOperacao_LostFocus()
sSQL = "SELECT CRT, ESTADO FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    vUFEmpresa = r("ESTADO")
    'vTipoCRT = r("CRT")
    If Left(cboDestOperacao.Text, 1) = 2 Then
        vAliqUFInter = Format(12, "#0.00")
        vAliqUFDest = Format(18, "#0.00")
    Else
        vAliqUFInter = Format(0, "#0.00")
        vAliqUFDest = Format(0, "#0.00")
    End If
End If

If cboDestOperacao.Text = "1 - Operação Interna" Then cboNatureza.Text = "5102"
If cboDestOperacao.Text = "2 - Operação Interestadual" Then cboNatureza.Text = "6102"
AtualizarCFOPCSTItens
End Sub


Private Sub cboFiltroNota_Click()
cboConNotaCliente.Visible = False
txtConNotaCodCliente.Visible = False
lblConNotaAno.Visible = False
cboConNotaAno.Visible = False
cboConNotaMes.Visible = False
mskConNotaInicial.Visible = False
mskConNotaFinal.Visible = False
cmdConNotaCal1.Visible = False
cmdConNotaCal2.Visible = False

If cboFiltroNota.Text = "TODAS" Then
    Exit Sub
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    lblConNotaAno.Caption = "Num. Nota:"
    lblConNotaAno.Visible = True
    cboConNotaCliente.Visible = True
    cboConNotaCliente.Text = ""
    cboConNotaCliente.SetFocus
ElseIf cboFiltroNota.Text = "CLIENTE" Then
    lblConNotaAno.Caption = "Cliente:"
    lblConNotaAno.Visible = True
    cboConNotaCliente.Visible = True
    cboConNotaCliente.Text = ""
    cboConNotaCliente.SetFocus
ElseIf cboFiltroNota.Text = "DATAS" Then
    lblConNotaAno.Caption = "Datas:"
    lblConNotaAno.Visible = True
    mskConNotaInicial.Visible = True
    mskConNotaFinal.Visible = True
    cmdConNotaCal1.Visible = True
    cmdConNotaCal2.Visible = True
    cmdConNotaCal1.SetFocus
ElseIf cboFiltroNota.Text = "MENSAL" Then
    lblConNotaAno.Caption = "Mês/Ano:"
    lblConNotaAno.Visible = True
    cboConNotaAno.Visible = True
    cboConNotaMes.Visible = True
'    If cboConNotaMes.Enabled = True Then cboConNotaMes.SetFocus
Else
End If
End Sub


Private Sub cboFiltroNota_GotFocus()
cboFiltroNota.Clear
cboFiltroNota.AddItem "TODAS"
cboFiltroNota.AddItem "NUM. NOTA"
cboFiltroNota.AddItem "CLIENTE"
cboFiltroNota.AddItem "DATAS"
cboFiltroNota.AddItem "MENSAL"

moCombo.AttachTo cboFiltroNota
End Sub


Private Sub cboFinalidade_GotFocus()
Dim VarText As String
VarText = cboFinalidade.Text

cboFinalidade.Clear
cboFinalidade.AddItem "1 - NFe NORMAL"
cboFinalidade.AddItem "2 - NFe COMPLEMENTAR"
cboFinalidade.AddItem "3 - NFe DE AJUSTE"
cboFinalidade.AddItem "4 - DEVOLUÇÃO/RETORNO"

If cboFinalidade.Text = "" Then cboFinalidade.Text = VarText
SelectControl cboFinalidade
moCombo.AttachTo cboFinalidade
End Sub


Private Sub cboFormaPgto_GotFocus()
Dim VarText As String
VarText = cboFormaPgto.Text

cboFormaPgto.Clear
cboFormaPgto.AddItem "01 = Dinheiro"
cboFormaPgto.AddItem "02 = Cheque"
cboFormaPgto.AddItem "03 = Cartão de Crédito"
cboFormaPgto.AddItem "04 = Cartão de Débito"
cboFormaPgto.AddItem "05 = Crédito Loja"
cboFormaPgto.AddItem "10 = Vale Alimentação"
cboFormaPgto.AddItem "11 = Vale Refeição"
cboFormaPgto.AddItem "12 = Vale Presente"
cboFormaPgto.AddItem "13 = Vale Combustível"
cboFormaPgto.AddItem "14 = Duplicata Mercantil"
cboFormaPgto.AddItem "15 = Boleto Bancário"
cboFormaPgto.AddItem "16 = Depósito Bancário"
cboFormaPgto.AddItem "18 = Transferência bancária"
cboFormaPgto.AddItem "19 = Programa de fidelidade"
cboFormaPgto.AddItem "20 = PIX"
cboFormaPgto.AddItem "90 = Sem pagamento"
cboFormaPgto.AddItem "99 = Outros"

If cboFormaPgto.Text = "" Then cboFormaPgto.Text = VarText
SelectControl cboFormaPgto
moCombo.AttachTo cboFormaPgto
End Sub


Private Sub cboIndicadorPagamento_Click()
cboIndicadorPagamento_LostFocus
End Sub

Private Sub cboIndicadorPagamento_LostFocus()
If cboIndicadorPagamento.Text = "0 - Pagamento à vista" Then
    frmDuplicata.Visible = False
ElseIf cboIndicadorPagamento.Text = "1 - Pagamento à prazo" Then
    frmDuplicata.Visible = True
    txtNumDup.Text = txtCodNota.Text
    txtTotalDup.Text = txtTotaldaNota.Text
    txtNumParcDup.Text = 1
    txtIntervaloDup.Text = 30
    'mskInicioDup.Text = mskEmissao.Text
    txtValorParcDup.Text = txtTotaldaNota.Text
    Calcular_Prazo
End If
End Sub


Private Sub cboIndicePedidos_Click()
lblInicialPedidos.Visible = False
mskInicialPedidos.Visible = False
cmdCalPedidos1.Visible = False
lblFinalPedidos.Visible = False
mskFinalPedidos.Visible = False
cmdCalPedidos2.Visible = False
lblClientePedidos.Visible = False
cboClientePedidos.Visible = False
lblConsCodPedido.Visible = False
txtConCodPedido.Visible = False
lblMesPedidos.Visible = False
cboMesPedidos.Visible = False
lblAnoPedidos.Visible = False
cboAnoPedidos.Visible = False

If cboIndicePedidos.Text = "PEDIDO" Then
    lblConsCodPedido.Visible = True
    txtConCodPedido.Visible = True
ElseIf cboIndicePedidos.Text = "CLIENTE" Then
    lblClientePedidos.Visible = True
    cboClientePedidos.Visible = True
ElseIf cboIndicePedidos.Text = "DATAS" Then
    lblInicialPedidos.Visible = True
    mskInicialPedidos.Visible = True
    cmdCalPedidos1.Visible = True
    lblFinalPedidos.Visible = True
    mskFinalPedidos.Visible = True
    cmdCalPedidos2.Visible = True
ElseIf cboIndicePedidos.Text = "MENSAL" Then
    lblMesPedidos.Visible = True
    cboMesPedidos.Visible = True
    lblAnoPedidos.Visible = True
    cboAnoPedidos.Visible = True
Else
End If
End Sub

Private Sub cboIndicePedidos_GotFocus()
cboIndicePedidos.Clear
cboIndicePedidos.AddItem "PEDIDO"
cboIndicePedidos.AddItem "CLIENTE"
cboIndicePedidos.AddItem "DATAS"
cboIndicePedidos.AddItem "MENSAL"

moCombo.AttachTo cboIndicePedidos
End Sub


Private Sub cboIndicePedidos_Validate(Cancel As Boolean)
cboIndicePedidos_Click
End Sub

Private Sub cboMesPedidos_GotFocus()
cboMesPedidos.Clear

cboMesPedidos.AddItem "Janeiro"
cboMesPedidos.AddItem "Fevereiro"
cboMesPedidos.AddItem "Março"
cboMesPedidos.AddItem "Abril"
cboMesPedidos.AddItem "Maio"
cboMesPedidos.AddItem "Junho"
cboMesPedidos.AddItem "Julho"
cboMesPedidos.AddItem "Agosto"
cboMesPedidos.AddItem "Setembro"
cboMesPedidos.AddItem "Outubro"
cboMesPedidos.AddItem "Novembro"
cboMesPedidos.AddItem "Dezembro"

moCombo.AttachTo cboMesPedidos
End Sub


Private Sub cboModFrete_GotFocus()
Dim VarText As String
VarText = cboModFrete.Text

cboModFrete.Clear
cboModFrete.AddItem "0 - Frete por conta do Remetente (CIF)"
cboModFrete.AddItem "1 - Frete por conta do Destinatário (FOB)"
cboModFrete.AddItem "2 - Frete por conta de Terceiros"
cboModFrete.AddItem "3 - Transporte Próprio por conta do Remetente"
cboModFrete.AddItem "4 - Transporte Próprio por conta do Destinatário"
cboModFrete.AddItem "9 - Sem Ocorrência de Transporte"

If cboModFrete.Text = "" Then cboModFrete.Text = VarText

'cboModFrete.AddItem ""
moCombo.AttachTo cboModFrete
End Sub


Private Sub cboNatureza_Change()
If cboNatureza.Text = "" Then Exit Sub

If Len(cboNatureza.Text) > 3 And Len(cboNatureza.Text) < 5 Then
    sSQL = "SELECT CodigoNatureza, NomeNatureza FROM NaturezaOperacaoNF where CodigoNatureza = " & cboNatureza.Text & " ORDER BY CodigoNatureza;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If r.BOF Then
        MsgBox "Natureza da operação incorreta!", vbInformation, "Aviso do Sistema"
        cboNatureza.Text = ""
        cboNatureza.SetFocus
        Exit Sub
    Else
        txtNatureza.Text = UCase(ValidateNull(r("NomeNatureza")))
    End If
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If
End Sub

Private Sub cboNatureza_Validate(Cancel As Boolean)
'preenche o combo naturaza
'Dim r As ADODB.Recordset

If cboNatureza.Text = "" Then Exit Sub

sSQL = "SELECT CodigoNatureza, NomeNatureza FROM NaturezaOperacaoNF where CodigoNatureza = " & cboNatureza.Text & " ORDER BY CodigoNatureza;"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    MsgBox "Natureza da operação incorreta!", vbInformation, "Aviso do Sistema"
    cboNatureza.Text = ""
    cboNatureza.SetFocus
    Exit Sub
Else
    txtNatureza.Text = UCase(ValidateNull(r("NomeNatureza")))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cboObservacao_GotFocus()
sSQL = "SELECT CodigoObservacao, Observacao FROM ObservacoesNFe;"
Set r = dbData.OpenRecordset(sSQL)
    
Do While Not r.EOF
   cboObservacao.AddItem r("Observacao")
   cboObservacao.ItemData(cboObservacao.NewIndex) = r("CodigoObservacao")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Sub cboObservacao_LostFocus()
On Error GoTo TrataErro

If cboObservacao.Text = "" Then txtCodOBS.Text = "": Exit Sub

If cboObservacao.ListIndex = -1 Then txtCodOBS.Text = "": Exit Sub

txtCodOBS = cboObservacao.ItemData(cboObservacao.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboTipoContribuinte_GotFocus()
Dim VarText As String
VarText = cboTipoContribuinte.Text

cboTipoContribuinte.Clear
cboTipoContribuinte.AddItem "1 - CONTRIBUINTE ICMS"
cboTipoContribuinte.AddItem "2 - CONTRIBUINTE ISENTO"
cboTipoContribuinte.AddItem "9 - NÃO CONTRIBUINTE"

If cboTipoContribuinte.Text = "" Then cboTipoContribuinte.Text = VarText
SelectControl cboTipoContribuinte
moCombo.AttachTo cboTipoContribuinte
End Sub

Private Sub cboTipoDest_GotFocus()
Dim VarText As String

VarText = cboTipoDest.Text

cboTipoDest.Clear
cboTipoDest.AddItem "CLIENTE"
cboTipoDest.AddItem "FORNECEDOR"

If cboTipoDest.Text = "" Then cboTipoDest.Text = VarText

SelectControl cboTipoDest
moCombo.AttachTo cboTipoDest
End Sub


Private Sub cboTipoNota_GotFocus()
Dim VarText As String
VarText = cboTipoNota.Text

cboTipoNota.Clear
cboTipoNota.AddItem "0 - ENTRADA"
cboTipoNota.AddItem "1 - SAÍDA"

If VarText <> "" Then cboTipoNota.Text = VarText
SelectControl cboTipoNota
moCombo.AttachTo cboTipoNota
End Sub





Private Sub cmdAdicionarItem_Click()
Dim vTotal As Double
'Dim totalRegistros As Long
'On Error GoTo erro
If txtCodNota.Text = "" Then Exit Sub
If txtCodProduto.Text = "" Then Exit Sub
If txtSubTotal.Text = "" Then Exit Sub
If Len(vNCM) < 8 Then ShowMsg "NCM INCORRETO!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbExclamation

sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

vgDb.BeginTrans

'insere os dados itens
Tb.AddNew
Load_Data_Itens
Tb.Update

vgDb.CommitTrans

'Call cmdRecalcularNF_Click   'desativei no dia do joelson

'Call DistribuirFrete
'Call DistribuirOutros
'Call DistribuirSeguro
'Call CalcularIPI
'Call CalcularDesconto
'Call AtualizarValorICMS
'Call MostrarValorProdutos
''Call MostrarValorBaseICMS
'If Left(cboDestOperacao.Text, 1) = "2" Then Call CalcularICMSInterItens
Call AtualizarTotaisNota

'Limpa_Tudo Me ' limpa tudo

'sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
'vTotal = SQLExecutaRetorno(sSQL, "r", 0)

'sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
'SQLExecuta sSQL

'EXIBIR NO GRID
Exibir_Itens

LimparObjetosProduto
LimparVariaveisItens

KeyCode = 0
TipoSelecaoConsulta = "0"
vTipoProduto = ""
cboDescricao.SetFocus
'cmdRecalcular_Click
Exit Sub
'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub
Private Sub SomarGridItens()
Dim Total As Currency
Dim SUBTOTAL As Currency
Dim Desc As Currency
Dim vICMS As Currency
Dim vIPI As Currency
Dim i As Integer

SUBTOTAL = 0
Desc = 0
Total = 0

'Sub-Total
With GridNotasItens
   For i = 1 To .rows - 1
      .Col = 0
      .Row = i
      
         .Col = 6
         SUBTOTAL = SUBTOTAL + (CDbl(.TextMatrix(.Row, 11)) * CDbl(.TextMatrix(.Row, 12)))
         Desc = Desc + .TextMatrix(.Row, 13)
         vICMS = vICMS + .TextMatrix(.Row, 10)
         vIPI = vIPI + .TextMatrix(.Row, 16)
         Total = Total + .TextMatrix(.Row, 14)
   Next
End With

'lblSubTotal.Caption = Format(CCur(SUBTOTAL), ocMONEY)
'lblSubTotal.Caption = Format(CCur(Total), ocMONEY)
'lblTotalDesc.Caption = Format(Desc, ocMONEY)
'lblValorNota.Caption = Format(Total, ocMONEY)

'txtTotaldosProdutos.Text = Format(SUBTOTAL, ocMONEY)
'txtTotaldosProdutos.Text = Format(Total, ocMONEY)
'txtBaseICMS.Text = Format(SUBTOTAL, ocMONEY)
'txtValorDesconto.Text = Format(Desc, ocMONEY)
'txtValorICMS.Text = Format(vICMS, ocMONEY)
'txtValorIPI.Text = Format(vIPI, ocMONEY)
'txtBaseICMS = Format(BaseICMS, ocMONEY)
'txtBaseICMS = Format(0, ocMONEY)
''MostrarValorBaseICMS

'CALCULAR OS TRIBUTOS
Dim varValorFrete As Currency
Dim varValorICMS As Currency
Dim varValorICMSST As Currency
Dim varValorIPI As Currency
Dim varValorDesconto As Currency
Dim varValorSeguro As Currency
Dim varValorOutrasDespesas As Currency
Dim varTotalNota As Currency

If txtValorFrete.Text = "" Then varValorFrete = 0 Else varValorFrete = txtValorFrete.Text
If txtValorICMS.Text = "" Then varValorICMS = 0 Else varValorICMS = txtValorICMS.Text
If txtValorICMSST.Text = "" Then varValorICMSST = 0 Else varValorICMSST = txtValorICMSST.Text
If txtValorIPI.Text = "" Then varValorIPI = 0 Else varValorIPI = txtValorIPI.Text
If txtValorDesconto.Text = "" Then varValorDesconto = 0 Else varValorDesconto = txtValorDesconto.Text
If txtValorSeguro.Text = "" Then varValorSeguro = 0 Else varValorSeguro = txtValorSeguro.Text
If txtValorOutrasDespesas.Text = "" Then varValorOutrasDespesas = 0 Else varValorOutrasDespesas = txtValorOutrasDespesas.Text

'varTotalNota = SUBTOTAL + varValorFrete + varValorICMS + varValorIPI + varValorSeguro + varValorOutrasDespesas
varTotalNota = SUBTOTAL
varTotalNota = varTotalNota - varValorDesconto
varTotalNota = varTotalNota + varValorFrete
varTotalNota = varTotalNota + txtValorIPI
txtTotaldaNota = FormatNumber(varTotalNota, 2)
'txtTotaldaNota = Format(Total, ocMONEY)
End Sub



Private Function Atualizar_Dados() As Boolean
'Comando de atualização
sSQL = "UPDATE setor SET setor = '" & cboSetor.Text & "' WHERE (cod_setor = " & txtCodigo.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados = dbData.Execute(sSQL)
End Function

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

mskEmissao = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub

Private Sub cmdCal2_Click()
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

mskSaida = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub


Private Sub cmdCalDuplic_Click()
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

mskInicioDup = Format(varData, "dd/mm/yy")   'Exibe a data no campo
'vDataFlexivel = True
'Calcular_Prazo
End Sub

Private Sub cmdCalPedidos1_Click()
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

mskInicialPedidos = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCalPedidos2_Click()
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

mskFinalPedidos = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdCancelarNota_Click()
Dim Justificativa As String
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))

 If MsgBox("Tem certeza que deseja Cancelar a Nota Fiscal?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
 Justificativa = InputBox("Informe a Justificativa do Cancelamento da Nota:", "Cancelamento da Nota", "DESISTENCIA DA COMPRA")
 vsNumeroNota = Val(vCodNota)
 Set TbNotas = Nothing
 RsOpen TbNotas, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, ChavedeAcesso, NumeroProtocolo,  " & _
                 "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                 "FROM NotaFiscal WHERE CodigoNota = " & Val(vCodNota)
 iRetorno = CancelaNFe(TbNotas("ChavedeAcesso"), TbNotas("NumeroProtocolo"), Justificativa, True)
 If iRetorno Then
    SQL = "UPDATE NotaFiscal SET " & _
          "Cancelada = 1, " & _
          "CanceladaProtocolo = " & NFeNumeroProtocolo & ", " & _
          "Justificativa = '" & Justificativa & "' " & _
          "WHERE CodigoNota = " & Val(vCodNota)
    vgDb.Execute SQL
    'RsOpen TbNotas, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & Val(vCodNota)
    'Load_Controls
    'FormatarGridNotas TbNotas
    Call cmdExibirConNotas_Click
 End If
End Sub

Private Sub CboCliente_GotFocus()
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboCliente.Text
codAtual = txtCodCliente.Text
cboCliente.Clear

If cboTipoDest.Text = "FORNECEDOR" Then
    sSQL = "SELECT DISTINCT RAZAO, codigo FROM FORNECEDOR ORDER BY razao;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboCliente.AddItem r("RAZAO")
       cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
       r.MoveNext
    Loop
Else
    sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       cboCliente.AddItem r("nome")
       cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
       r.MoveNext
    Loop
End If

If r.State <> 0 Then r.Close
Set r = Nothing

cboCliente.Text = itemAtual
txtCodCliente.Text = codAtual
SelectControl cboCliente
moCombo.AttachTo cboCliente
End Sub
Private Sub CboCliente_LostFocus()
On Error GoTo TrataErro

If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub

'If cmdAlterar.Enabled = False Then
If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
'End If

txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboDescricao_GotFocus()
Dim vNomeProduto As String

'moCombo.AttachTo cboDescricao

'If TipoSelecaoConsulta = "0" Or TipoSelecaoConsulta = "2" Then

    vNomeProduto = cboDescricao.Text
    'If cboDescricao.ListIndex = -1 Then
        cboDescricao.Clear
        
'        If Len(cboDescricao.Text) > 3 Then
            Dim vUltimoValorVenda As String     '===================TER QUE COLOCAR DEPOIS PARA TODOS OS TIPOS DE VENDAS
            vUltimoValorVenda = " (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) "
                 
            'carrega a consulta
            sSQL = "SELECT DISTINCT top 200 descricao, codigo " & _
            "FROM produtos WHERE (descricao LIKE '%" & cboDescricao.Text & "%') AND (produtos.ativo = 1) and " & vUltimoValorVenda & " > 0  " & _
            "ORDER BY descricao;"
            Set r = dbData.OpenRecordset(sSQL)
            
            Do While Not r.EOF
               cboDescricao.AddItem ValidateNull(r("descricao"))
                cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
               r.MoveNext
            Loop
       ' End If
    'End If
    
    cboDescricao.Text = vNomeProduto
    SelectControl cboDescricao
    moCombo.AttachTo cboDescricao
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
'End If

End Sub


Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboDescricao_LostFocus()
If TipoSelecaoConsulta = "0" Or TipoSelecaoConsulta = "2" Then

If cboDescricao.Text = "" Then
    LimparObjetosProduto
    TipoSelecaoConsulta = "0"
Else
    If cboDescricao.ListIndex = -1 Then
        ShowMsg "Produto não cadastrado.", vbExclamation
        LimparObjetosProduto
        TipoSelecaoConsulta = "0"
        cboDescricao.SetFocus
        'Exit Sub
    Else
        txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)
        TipoSelecaoConsulta = "2"
        MostrarValorVenda
        Mostrar_Aliquotas_Produto
        txtQuant.SetFocus
    End If
End If

If TipoSelecaoConsulta = "1" Then
    txtCodBarra.BackColor = &HC0FFFF
    cboDescricao.BackColor = &HFFFFFF
    cboDescricao.Locked = True
ElseIf TipoSelecaoConsulta = "2" Then
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HC0FFFF
    txtCodBarra.Locked = True
Else
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HFFFFFF
    txtCodBarra.Locked = False
    cboDescricao.Locked = False
End If

    
    'If cboDescricao.ListIndex = -1 Then
    '    txtCodProduto.Text = ""
    '    TipoSelecaoConsulta= "0"
    '    txtCodBarra.Locked = False
    '    cboDescricao.Text = ""
    '    txtCodBarra.Text = ""
    '    Exit Sub
    'End If


   'txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)
   
   'If txtCodProduto.Text = "" Then Exit Sub
   
   
   
    'If txtCodProduto.Text = "" Then
    '    TipoSelecaoConsulta= "0"
    '    txtCodBarra.Locked = False
    '    cboDescricao.Text = ""
    '    txtCodBarra.Text = ""
    '    Exit Sub
    'End If
   
    'Mostrar_Aliquotas_Produto

   'If r.BOF Then ShowMsg "Produto não cadastrado.", vbExclamation
   
End If
End Sub
Private Sub Mostrar_Aliquotas_Produto()
If txtCodProduto.Text = "" Then Exit Sub
sSQL = "SELECT codigo, descricao, INF_ADICIONA, EAN, COD_BARRA, unid_medida, ncm, tamanho, REF, fabricante, CFOP, ICMSCST, ICMSAliq, pRedBC, modBC, piscst, pisAliq, cofinscst, cofinsAliq, ipicst, ipiAliq, cest, pMVAST, pICMSST, pRedBCST, CASE WHEN abs(combustivel) = 1 THEN 'Combustível' ELSE '' END as vTProduto FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

 If Not r.BOF Then
    If TipoSelecaoConsulta = "2" Then
        txtCodBarra.Text = r("COD_BARRA")
    ElseIf TipoSelecaoConsulta = "1" Then
        If tipoEmpresa = 4 Then
            cboDescricao.Text = ValidateNull(r("descricao")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante")) & " /  " & r("REF")
        Else
            cboDescricao.Text = ValidateNull(r("descricao"))
        End If
    End If
    
     vInfAdd = ValidateNull(r("INF_ADICIONA"))
     vEAN = ValidateNull(r("EAN"))
     vUnid_medida = Left(ValidateNull(r("unid_medida")), 2)
     vCFOP = ValidateNull(r("CFOP"))
     vNCM = ValidateNull(r("ncm"))
     vICMSCST = ValidateNull(r("ICMSCST"))
     vICMSAliq = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
     vpRedBC = FormatNumber(ValidateNull(r("pRedBC")), 4)
     vPISCST = ValidateNull(r("piscst"))
     vPISALIQ = Format(ValidateNull(r("pisAliq")), "##,##0.00")
     vCOFINSCST = ValidateNull(r("cofinscst"))
     vCOFINSALIQ = Format(ValidateNull(r("cofinsAliq")), "##,##0.00")
     vIPICST = ValidateNull(r("ipicst"))
     vIPIALIQ = Format(ValidateNull(r("ipiAliq")), "##,##0.00")
     vCEST = ValidateNull(r("cest"))
     vModBC = ValidateNull(r("modBC"))
     vPMVAST = Format(ValidateNull(r("pMVAST")), "##,##0.00")
     vPICMSST = Format(ValidateNull(r("pICMSST")), "##,##0.00")
     vPRedBCST = Format(ValidateNull(r("pRedBCST")), "##,##0.00")

    'If CBool(r("combustivel")) = True Then
        vTipoProduto = r("vTProduto")
     'Else
     '   vTipoProduto = ""
     'End If
     
 Else
     ShowMsg "Produto não cadastrado.", vbExclamation
     TipoSelecaoConsulta = "0"
     cboDescricao.Text = ""
     txtCodBarra.Text = ""
    LimparObjetosProduto
     
     vEAN = ""
     vInfAdd = ""
     vUnid_medida = ""
     vCFOP = ""
     vNCM = ""
     vICMSCST = ""
     vICMSAliq = ""
     vModBC = ""
     vPMVAST = ""
     vPICMSST = ""
     vPRedBCST = ""
     vpRedBC = ""
     vPISCST = ""
     vPISALIQ = ""
     vCOFINSCST = ""
     vCOFINSALIQ = ""
     vIPICST = ""
     vIPIALIQ = ""
     vCEST = ""
     vTipoProduto = ""
 End If
 If r.State <> 0 Then r.Close
End Sub

Private Sub MostrarValorVenda()
Dim vrVenda As Currency
If txtCodProduto.Text = "" Then Exit Sub

'mostrar o ultimo preço de compra
sSQL = "SELECT TOP 1 VALOR_VV FROM Produtos_Precos WHERE (COD_PRODUTO = " & txtCodProduto & ") ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then vrVenda = r("VALOR_VV")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValor.Text = Format(vrVenda, ocMONEY)
End Sub

Private Sub cboNatureza_GotFocus()
Dim r As ADODB.Recordset
Dim itemAtual As String

itemAtual = cboNatureza.Text
cboNatureza.Clear

sSQL = "SELECT CodigoNatureza, NomeNatureza FROM NaturezaOperacaoNF ORDER BY CodigoNatureza;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboNatureza.AddItem r("CodigoNatureza")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboNatureza.Text = itemAtual
SelectControl cboNatureza
moCombo.AttachTo cboNatureza
End Sub

Private Sub cboTransporte_GotFocus()
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboTransporte.Text
codAtual = txtCodTransporte.Text
cboTransporte.Clear

sSQL = "SELECT DISTINCT fantasia, codigo FROM transportadora ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboTransporte.AddItem r("fantasia")
   cboTransporte.ItemData(cboTransporte.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboTransporte.Text = itemAtual
txtCodTransporte.Text = codAtual
moCombo.AttachTo cboTransporte
End Sub


Private Sub cboTransporte_LostFocus()
On Error GoTo TrataErro

If cboTransporte.Text = "" Then txtCodTransporte.Text = "": Exit Sub
If cboTransporte.ListIndex = -1 Then txtCodTransporte.Text = "": Exit Sub

txtCodTransporte = cboTransporte.ItemData(cboTransporte.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cmdCancelar_Click()

If vTipoEdicaoNFe = "Novo" Then
    ' Nota criada pelo cmdNovo mas nao salva -- confirma exclusao
    If MsgBox("Deseja cancelar a nota em digitação? Os dados serão excluídos.", vbQuestion + vbYesNo, "Online Commerce") <> vbYes Then Exit Sub
    If txtCodNota.Text <> "" Then
        SQLExecuta "DELETE FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
        SQLExecuta "DELETE FROM NotaFiscal WHERE CodigoNota = " & Val(txtCodNota.Text)
    End If
End If

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False

frmNota.Enabled = False
frmDestinatario.Enabled = False
Tab_Totais.Enabled = False
Tab_Produtos.Enabled = False
frmItens.Enabled = False
TipoSelecaoConsulta = "0"
Tab_Totais.Tab = 0
Tab_Produtos.Tab = 0

If TbNotas.EditMode <> 0 Then TbNotas.CancelUpdate

LimparObjetosNota
LimparObjetosDestinatario
LimparObjetosProduto
LimparObjetosNotaTotais
LimparObjestosNotaOutros
LimparGridItensNota
vTipoEdicaoNFe = ""

End Sub

Private Sub cmdCartaCorrecao_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
MostrarCorrecao
frmCorreção.Visible = True
txtCorrecao.SetFocus

'Tab_Produtos.Tab = 6
'Tab_Produtos.Enabled = True
End Sub


Private Sub cmdCCeExcluir_Click()
i = Grid_Correcao.Row
If Grid_Correcao.TextMatrix(i, 1) = "" Then Exit Sub
If Grid_Correcao.TextMatrix(i, 8) = "ENVIADO" Then MsgBox "A Carta de Correção já foi transmitida!", vbInformation, "Aviso do Sistema": Exit Sub
Dim idCartaCorrecao As Long
idCartaCorrecao = Grid_Correcao.TextMatrix(i, 1)
dbData.Execute "DELETE FROM NFeCartaCorrecao WHERE CodigoCartaCorrecao = " & Val(idCartaCorrecao)
MostrarCorrecao
cmdCCeImprimir.Enabled = False
cmdCCeTransmitir.Enabled = False
cmdCCeExcluir.Enabled = False
End Sub

Private Sub cmdCCeImprimir_Click()
Dim ComandoSQL As String, SeqEvento As Integer, textoCorrecao As String
Dim idCartaCorrecao As Long
Dim vChave As String
Dim vDataProt As Date
vChave = GridNotas.TextMatrix(GridNotas.Row, 8)
vDataProt = GridNotas.TextMatrix(GridNotas.Row, 11)

Dim objNFe As New snfe.Util

   On Error GoTo deuErro
   
   Set objNFe = New snfe.Util
   
   i = Grid_Correcao.Row
      
   If Grid_Correcao.TextMatrix(i, 1) = "" Then Exit Sub
   If Grid_Correcao.TextMatrix(i, 8) <> "ENVIADO" Then MsgBox "A Carta de Correção ainda não foi transmitida!", vbInformation, "Aviso do Sistema": Exit Sub

   idCartaCorrecao = Grid_Correcao.TextMatrix(i, 1)
   
   If idCartaCorrecao = 0 Then Exit Sub
   
   SeqEvento = Grid_Correcao.TextMatrix(i, 4)

   dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.path)
   xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(vDataProt, "yyyymm") & "\" & vChave & "-procNFe.xml"
   xCaminhoXMLAuxiliar = dirXML & "\nfe\arquivos\procEventoNFe\110110" & vChave & LPad(SeqEvento, 2, "0") & "-procEventoNFe.xml"
   xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\CCe" & vChave & "_" & LPad(SeqEvento, 2, "0") & ".pdf"
   
   iRetorno = ConfiguraDLLNFeNFCe(55, "1", objNFe)

   Call objNFe.EventoImprimir(xCaminhoXML, xCaminhoXMLAuxiliar, False, "", True, xCaminhoPDF, False, "", True)
   
   Set objNFe = Nothing
   
   Exit Sub

deuErro:
   MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
   Err.Clear
   Set objNFe = Nothing
End Sub


Private Sub cmdCCeSalvar_Click()
If txtCorrecao.Text = "" Then Exit Sub

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

txtCorrecao.Text = ""

MostrarCorrecao
End Sub


Private Function Inserir_Dados() As Boolean
vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))

'autonumeração
Dim vNovoCodigo As Integer
sSQL = "SELECT MAX(CodigoCartaCorrecao) r FROM NFeCartaCorrecao "
vNovoCodigo = SQLExecutaRetorno(sSQL, "r", 0) + 1

'autonumeração
Dim vNovoEvento As Integer
sSQL = "SELECT MAX(SeqCorrecao) r FROM NFeCartaCorrecao where CodigoNota= " & vCodNota & " "
vNovoEvento = SQLExecutaRetorno(sSQL, "r", 0) + 1

'Comando de inclusão
sSQL = "INSERT INTO NFeCartaCorrecao (" & _
   "CodigoCartaCorrecao, CodigoNota, Data, SeqCorrecao, TextoCorrecao, NumeroProtocolo, Enviada) VALUES (" & _
   vNovoCodigo & ", " & vCodNota & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & vNovoEvento & ", '" & txtCorrecao.Text & "', 0, 0)"

'Retorna o resultado da atualização
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub cmdCCeTransmitir_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))
Dim vChave As String
vChave = GridNotas.TextMatrix(GridNotas.Row, 8)

Dim ComandoSQL As String, Parametros As New ADODB.Recordset, SeqEvento As Integer, textoCorrecao As String
Dim idCartaCorrecao As Long
Dim objNFe As New snfe.Util

   On Error GoTo deuErro
   
   i = Grid_Correcao.Row
   If Grid_Correcao.TextMatrix(i, 1) = "" Then Exit Sub
   If Grid_Correcao.TextMatrix(i, 8) = "ENVIADO" Then MsgBox "A Carta de Correção já foi transmitida!", vbInformation, "Aviso do Sistema": Exit Sub

   idCartaCorrecao = Grid_Correcao.TextMatrix(i, 1)
   
   If idCartaCorrecao = 0 Then Exit Sub
   
   'txtCodNota 'onde fica o codigo da nota

   Set objNFe = New snfe.Util

   vsSQL = "SELECT * FROM Empresa"
   RsOpen Parametros, vsSQL
    
   iRetorno = ConfiguraDLLNFeNFCe(55, "1", objNFe)
    
   SeqEvento = Grid_Correcao.TextMatrix(i, 4)
    
   textoCorrecao = Grid_Correcao.TextMatrix(i, 5)
    
    iRetorno = objNFe.CartaCorrecao(Parametros!CNPJ, CLng(vCodNota), SeqEvento, vChave, textoCorrecao, xCaminhoXMLAuxiliar)
    
    If Not iRetorno Then GoTo Caifora
    cStat = objNFe.retEnvEvento.cStat
    NFeMotivo = objNFe.retEnvEvento.xMotivo
    If cStat = 128 Then
       cStat2 = objNFe.retEnvEvento.retEvento.infEvento.cStat
       NFeValidate = objNFe.retEnvEvento.retEvento.infEvento.xMotivo
       NFeNumeroProtocolo = objNFe.retEnvEvento.retEvento.infEvento.nProt
       NFeDataHora = objNFe.retEnvEvento.retEvento.infEvento.dhRegEvento
    Else
       cStat2 = objNFe.retEnvEvento.retEvento.infEvento.cStat
       NFeValidate = objNFe.retEnvEvento.retEvento.infEvento.xMotivo
       NFeNumeroProtocolo = ""
       NFeDataHora = objNFe.retEnvEvento.retEvento.infEvento.dhRegEvento
    End If
  
    If cStat2 = 135 Then
       GoTo continua
    Else
       If cStat2 > 0 Then
          MsgBox str(cStat2) & " - " & NFeValidate, vbInformation, "ERRO"
       Else
          MsgBox str(cStat) & " - " & NFeMotivo, vbInformation, "ERRO"
       End If
       GoTo Caifora
    End If
         
continua:
    msgResultado = "Protocolo.: " + NFeNumeroProtocolo & vbCrLf
    msgResultado = msgResultado + "Data/Hora: " & NFeDataHora & vbCrLf
    msgResultado = msgResultado + "Resposta da Fazenda.: " + str(cStat2) & " - " & NFeValidate
    
    MsgBox msgResultado, vbInformation + vbOKOnly, "Envio CCe"
    
    If cStat2 = 135 Then
       ComandoSQL = "UPDATE NFeCartaCorrecao SET " & _
                    "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
                    "DataHoraProcotolo = '" & NFeDataHora & "', " & _
                    "Enviada = 1 " & _
                    "WHERE CodigoCartaCorrecao = " & idCartaCorrecao
       vgDb.Execute ComandoSQL
    End If
  
    Screen.MousePointer = vbDefault
    Set Parametros = Nothing
    Set objNFe = Nothing
    MostrarCorrecao
    cmdCCeImprimir.Enabled = False
    cmdCCeTransmitir.Enabled = False
    cmdCCeExcluir.Enabled = False
    Exit Sub

Caifora:
    Set Parametros = Nothing
    Set objNFe = Nothing
    If Vazio(NFeMotivo) Then MsgBox NFeResposta & vbNewLine & "ERRO NO ENVIO DA CARTA DE CORREÇÃO", vbCritical, vgAtencao
    Screen.MousePointer = vbDefault
    Exit Sub

Resume
deuErro:
    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO"
    Err.Clear
    Screen.MousePointer = vbDefault
    Set Parametros = Nothing
    Set objNFe = Nothing
    Exit Sub

'(CodigoCartaCorrecao = " & Grid_Correcao.TextMatrix(i, 1) & ")
'(CodigoNota = " & Grid_Correcao.TextMatrix(i, 2) & ")
'(Data = " & Grid_Correcao.TextMatrix(i, 3) & ")
'(SeqCorrecao = " & Grid_Correcao.TextMatrix(i, 4) & ")
'(TextoCorrecao = " & Grid_Correcao.TextMatrix(i, 5) & ")
'(NumeroProtocolo = " & Grid_Correcao.TextMatrix(i, 6) & ")
'(DataHoraProcotolo = " & Grid_Correcao.TextMatrix(i, 7) & ")

End Sub

Private Sub cmdConNotaCal1_Click()
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

mskConNotaInicial = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConNotaCal2_Click()
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

mskConNotaFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConsultaNCMean_Click()
Dim varNomeProduto As String
varNomeProduto = GridNotasItens.TextMatrix(GridNotasItens.Row, 2)
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultar_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub

vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))

'If ((GridNotas.TextMatrix(GridNotas.Row, 9)) = Empty) Or ((GridNotas.TextMatrix(GridNotas.Row, 9)) = "") Then Exit Sub
If ((GridNotas.TextMatrix(GridNotas.Row, 8)) = "") And ((GridNotas.TextMatrix(GridNotas.Row, 9)) = "0") Then Exit Sub
vsNumeroNota = Val(vCodNota)
If ((GridNotas.TextMatrix(GridNotas.Row, 9) = Empty) Or ((GridNotas.TextMatrix(GridNotas.Row, 9) = "")) And (GridNotas.TextMatrix(GridNotas.Row, 10)) = "0") Then
    Call consultaNFe(GridNotas.TextMatrix(GridNotas.Row, 8), False)
ElseIf (GridNotas.TextMatrix(GridNotas.Row, 8) <> Empty) Or ((GridNotas.TextMatrix(GridNotas.Row, 8) <> "")) Then
    Call consultaNFe(GridNotas.TextMatrix(GridNotas.Row, 8), False)
Else
   Exit Sub
End If
''If (((GridNotas.TextMatrix(GridNotas.Row, 9)) <> Empty) Or ((GridNotas.TextMatrix(GridNotas.Row, 9)) <> "0")) And (Text31.Text = Empty Or Text31.Text = "0") Then
''   vsNumeroNota = Val(txtCodNota.Text)
''   ConsultaRecibo (GridNotas.TextMatrix(GridNotas.Row, 9)), (GridNotas.TextMatrix(GridNotas.Row, 8)), "1", True
''Else
'   consultaNFe (GridNotas.TextMatrix(GridNotas.Row, 8))
'   If cStat = 100 Or cStat = 150 Then
'      SQL = "UPDATE NotaFiscal SET " & _
'            "Enviada = 1, " & _
'            "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
'            "DataHoraProcotolo = '" & NFeDataHora & "' " & _
'            "WHERE CodigoNota = " & Val(vCodNota)
'      vgDb.Execute SQL
'      'SQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
'      '      "(" & Val(txtCodNota.Text) & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
'      'vgDb.Execute SQL
'   ElseIf cStat = 101 Then
'      SQL = "UPDATE NotaFiscal SET " & _
'            "Cancelada = 1, " & _
'            "CanceladaProtocolo = " & NFeNumeroProtocolo & " " & _
'            "WHERE CodigoNota = " & Val(vCodNota)
'      vgDb.Execute SQL
'   ElseIf cStat = 110 Then
'      SQL = "UPDATE NotaFiscal SET " & _
'            "Denegada = 1, " & _
'            "NumeroProtocolo = " & NFeNumeroProtocolo & " " & _
'            "WHERE CodigoNota = " & Val(vCodNota)
'      vgDb.Execute SQL
'   End If
''End If
''RsOpen TbNotas, "SELECT *,  " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 0 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) AS Status " & _
                "FROM NotaFiscal WHERE CodigoNota = " & Val(vCodNota)
''Load_Controls
''FormatarGridNotas TbNotas
If cStat = 100 Or cStat = 150 Then
      SQL = "UPDATE NotaFiscal SET " & _
            "Enviada = 1, " & _
            "EmProcessamento = 0, " & _
            "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
            "DataHoraProcotolo = '" & NFeDataHora & "' " & _
            "WHERE CodigoNota = " & Val(vCodNota)
      vgDb.Execute SQL
      'SQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
      '      "(" & Val(txtCodNota.Text) & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
      'vgDb.Execute SQL
   ElseIf cStat = 101 Then
      SQL = "UPDATE NotaFiscal SET " & _
            "Cancelada = 1, " & _
            "CanceladaProtocolo = " & NFeNumeroProtocolo & " " & _
            "WHERE CodigoNota = " & Val(vCodNota)
      vgDb.Execute SQL
   ElseIf cStat = 110 Then
      SQL = "UPDATE NotaFiscal SET " & _
            "Denegada = 1, " & _
            "NumeroProtocolo = " & NFeNumeroProtocolo & " " & _
            "WHERE CodigoNota = " & Val(vCodNota)
      vgDb.Execute SQL
   Else
      SQL = "UPDATE NotaFiscal SET " & _
            "Enviada = 0 " & _
            "WHERE CodigoNota = " & Val(vCodNota)
      vgDb.Execute SQL
   End If
   
   If Not Vazio(NFeChaveAcesso) Then
      Clipboard.Clear
      Clipboard.SetText NFeChaveAcesso
   End If
   
Call cmdExibirConNotas_Click
End Sub

Private Sub cmdConsultarCest_Click()
Dim varNomeProduto As String
varNomeProduto = GridNotasItens.TextMatrix(GridNotasItens.Row, 6)
ShellExecute hwnd, "open", "http://www.buscacest.com.br/?utf8=" + Chr(95) + "&ncm=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
'http://www.buscacest.com.br/?utf8=%E2%9C%93&ncm=1704.90.20
End Sub

Private Sub cmdConsultarCliente_Click()
If txtCodCliente.Text = "" Then MsgBox "Escolha um cliente!", vbInformation, "Aviso do Sistema": Exit Sub
Dim varCodCliente As String
varCodCliente = txtCodCliente.Text

If cboTipoDest.Text = "CLIENTE" Then
    If ShowMsg("Deseja atualizar o cliente " & cboCliente.Text & " ?", vbInformation + vbYesNo) = vbYes Then
        Load Clientes_Cadastro
        Clientes_Cadastro.SSTab1.Tab = 0
        Clientes_Cadastro.cmdNovo.Enabled = False
        Clientes_Cadastro.cmdSalvar.Enabled = False
        Clientes_Cadastro.cmdCancelar.Enabled = False
        Clientes_Cadastro.txtCodigo.Text = varCodCliente
        Clientes_Cadastro.Show 1
    End If
ElseIf cboTipoDest.Text = "FORNECEDOR" Then
    If ShowMsg("Deseja atualizar o fornecedor " & cboCliente.Text & " ?", vbInformation + vbYesNo) = vbYes Then
        Load Fornecedor_Cadastro
        Fornecedor_Cadastro.SSTab1.Tab = 0
        Fornecedor_Cadastro.cmdNovo.Enabled = False
        Fornecedor_Cadastro.cmdAlterar.Enabled = True
        Fornecedor_Cadastro.frm_Principal.Enabled = True
        Fornecedor_Cadastro.cmdCancelar.Enabled = False
        Fornecedor_Cadastro.txtCodigo.Text = varCodCliente
        Fornecedor_Cadastro.Show 1
    End If
End If
End Sub

Private Sub cmdConsultarNCM_Click()
Dim varNomeProduto As String
varNomeProduto = Replace(GridNotasItens.TextMatrix(GridNotasItens.Row, 4), " ", "+")
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultarProduto_Click()
If GridNotasItens.rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim varCodProduto As String
varCodProduto = GridNotasItens.TextMatrix(GridNotasItens.Row, 3)
'MsgBox GridNotasItens.TextMatrix(GridNotasItens.Row, 3)
If ShowMsg("Deseja atualizar o produto " & GridNotasItens.TextMatrix(GridNotasItens.Row, 4) & " ?", vbInformation + vbYesNo) = vbYes Then

Load Produtos_Cadastro
Produtos_Cadastro.SSTab1.Tab = 0
Produtos_Cadastro.cmdNovo.Enabled = False
Produtos_Cadastro.cmdSalvar.Enabled = False
Produtos_Cadastro.cmdCancelar.Enabled = False
vTipoEdicao = "Edicao"
Produtos_Cadastro.txtCodigo.Text = varCodProduto
Produtos_Cadastro.Show 1
End If
End Sub

Private Sub cmdConverterNFe_Click()
If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub
If GridPedidos.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
If GridPedidos.TextMatrix(GridPedidos.Row, 2) = "SIM" Then MsgBox "Esse pedido já foi transformado em NFe!", vbInformation, "Online Commerce": Exit Sub
If ShowMsg("Deseja realmente transformar o pedido: " & GridPedidos.TextMatrix(GridPedidos.Row, 1) & " em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

txtCodPedido.Text = (GridPedidos.TextMatrix(GridPedidos.Row, 1))

vTipoEdicaoNFe = "Edicao"
picAguarde2.Visible = True
GravarPedido
txtInfComple.Text = "NFe referente a venda Nº " & txtCodPedido.Text
Frm_NF.Tab = 0
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
frmNota.Enabled = True
frmDestinatario.Enabled = True
frmItens.Enabled = True
Tab_Totais.Enabled = True
Tab_Produtos.Enabled = True
cmdRecalcular_Click
cmdExibirPedidos_Click
picAguarde2.Visible = False
End Sub

Private Sub cmdCopiarChave_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
'pegar o codigo da NFe Original
Dim vChave As String

vChave = GridNotas.TextMatrix(GridNotas.Row, 8)
Clipboard.Clear
Clipboard.SetText vChave
MsgBox "Chave de acesso copiada com sucesso!", vbInformation, "Aviso do Sistema"
End Sub

Private Sub cmdCriarDuplicata_Click()
If txtCodNota.Text = "" Then Exit Sub
If txtNumDup.Text = "" Then Exit Sub
If txtTotalDup.Text = "" Then Exit Sub
If txtIntervaloDup.Text = "" Then Exit Sub
If mskInicioDup.Text = "" Then Exit Sub
If txtValorParcDup.Text = "" Then Exit Sub
If txtNumParcDup.Text = "" Then Exit Sub
If cboFormaPgto.Text = "" Then MsgBox "Escolha uma forma de pagamento!", vbInformation, "Aviso do Sistema": cboFormaPgto.SetFocus: Exit Sub

'verificar se já existe duplicata criada
sSQL = "SELECT * FROM NotaFiscalParcelas WHERE CodigoNota = " & Val(txtCodNota.Text)
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    MsgBox "Já existe duplicatas criadas para essa nota fiscal." + vbCrLf + "Caso deseja mudar as duplicatas, apague primeiro as anteriores criadas!", vbInformation, "Aviso do Sistema"
    Exit Sub
Else
    'parcelas
    vVencimento = CDate(mskInicioDup.Text)
    vNumParc = 1
    
    CalcularParcelas CCur(txtTotalDup), CInt(txtNumParcDup), arrayParc
    
    'criar as parcelas Left(cboFormaPgto.Text, 2)
    For i = 1 To CInt(txtNumParcDup)
       dbData.Execute "INSERT INTO NotaFiscalParcelas (CodigoNota, Sequencia, Documento, CodigoFormaPagamento, Vencimento, ValorDocumento) VALUES (" & _
          txtCodNota.Text & ",  " & vNumParc & ", " & txtNumDup.Text & ", " & Left(cboFormaPgto.Text, 2) & ", '" & Format$(vVencimento, "yyyy-dd-MM") & "', " & _
          Replace(arrayParc(i), ",", ".") & ");"
       
        If txtIntervaloDup.Text = "30" Then
            vVencimento = Format(DateAdd("m", Val(1), vVencimento), "dd/mm/yy")
        Else
            vVencimento = Format(DateAdd("d", Val(txtIntervaloDup.Text), vVencimento), "dd/mm/yy")
        End If
       
       vNumParc = vNumParc + 1
    Next
End If

If r.State <> 0 Then r.Close
Set r = Nothing

Exibir_Duplicatas

LimparObjetosDuplicata
End Sub
Private Sub cmdDuplicar_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub

'pegar o codigo da NFe Original
Dim varCodNotaOrigem As Integer
Dim varNumNotaOrigem As Integer

vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))
vTipoEdicaoNFe = "Edicao"
varCodNotaOrigem = vCodNota
varNumNotaOrigem = (GridNotas.TextMatrix(GridNotas.Row, 2))

'On Error GoTo ErrLoad

'autonumeração
Dim ConsultaSQL As String
Dim tbNota As ADODB.Recordset

ConsultaSQL = "SELECT ISNULL(MAX(numeronota), 0) AS Maior_nota FROM NotaFiscal"
Set tbNota = dbData.OpenRecordset(ConsultaSQL)
      
'preecher objetos do form
Dim totalRegistros As Long
RsOpen TbNotas, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, serienf,  " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM NotaFiscal"
                
If TbNotas.RecordCount > 0 Then totalRegistros = TbNotas.RecordCount

    LimparObjetosNota
    LimparObjetosDestinatario
    LimparObjetosProduto
    LimparObjetosNotaTotais
    LimparObjestosNotaOutros
    LimparGridItensNota
    
    txtCodPedido.Text = 0
    
    If TbNotas.EOF And TbNotas.BOF Then
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = "1"
        txtSerie.Text = "1"
    Else
        TbNotas.MoveLast
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = TbNotas("CodigoNota") + 1
        txtSerie.Text = TbNotas("serienf")
    End If
    
    'preencher os objetos
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
    
    cboTipoNota.Text = "1 - SAÍDA"
    cboTipoContribuinte.Text = "9 - NÃO CONTRIBUINTE"
    cboConsumidorFinal.Text = "1 - SIM"
    
    cboModFrete.Text = ""
    
    txtValorFrete.Text = "0,00"
    txtValorOutrasDespesas.Text = "0,00"
    txtVolPesoBruto.Text = "0,00"
    txtVolPesoLiquido.Text = "0,00"
    txtValorSeguro.Text = "0,00"
    txtBaseICMSST.Text = "0,00"
    txtBaseICMS.Text = "0,00"
    txtValorICMS.Text = "0,00"
    txtBaseICMSST.Text = "0,00"
    txtValorICMSST.Text = "0,00"
    txtValorIPI.Text = "0,00"
    txtValorDesconto.Text = "0,00"
    
    cboCliente = "CONSUMIDOR"
    txtCodCliente = "1"

    cboDestOperacao.Text = "1 - Operação Interna"
    'txtInfAdicionais = Format(rTabela("InformacoesAdicionais"), "@")
    cboNatureza = "5102"
    txtNatureza = "VENDA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS"
    cboFinalidade = "1 - NFe NORMAL"
    cboTipoDest = "CLIENTE"
    mskEmissao = Format(Date, "dd/mm/yyyy")
    mskSaida = Format(Date, "dd/mm/yyyy")
    mskHora = Format(Time(), "HH:MM:ss")
    
    txtVolPesoBruto = Format(0, "@")
    txtVolPesoLiquido = Format(0, "@")
    'txtPlacaUF = Format(0, "@")
    cboModFrete = "9 - SEM FRETE"
    'txtCodTransporte = Format(0, "@")
    'cboTransporte = Format(0, "@")
    'txtPlaca = Format(0, "@")
    'txtVolQuant = Format(0, "@")
    'txtVolEspecie = Format(0, "@")
    'txtVolMarca = Format(0, "@")
    'txtVolNumeracao = Format(0, "@")
    'txtCodObservacao = Format(0, "@")
    
    txtValorSeguro = Format(0, "##,##0.00")
    txtValorOutrasDespesas = Format(0, "##,##0.00")
    txtValorFrete = Format(0, "##,##0.00")
    txtBaseICMS = Format(0, "##,##0.00")
    txtBaseICMSST = Format(0, "##,##0.00")
    txtValorIPI = Format(0, "##,##0.00")
    txtValorICMS = Format(0, "##,##0.00")
    txtValorICMSST = Format(0, "##,##0.00")
    txtValorDesconto = Format(0, "##,##0.00")
    
    'transmissão
    'Text30 = Format(0, "@")
    'Text31 = Format(0, "@")
    'Text32 = Format(0, "@")
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
    
    txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
    txtCodPedido.Text = 0

    LimparObjetosProduto
    
    'salva nota
    RsOpen TbNotas, "SELECT * FROM NotaFiscal"
    
    TbNotas.AddNew

    
'preecher os itens==========================================================

Dim tblItensPedido As ADODB.Recordset

'Atualiza a base de dados (funcionando)
Dim VarCodNota As Integer
VarCodNota = CInt(txtCodNota.Text)

sSQL = "INSERT INTO NotaFiscalItens ( " & _
        "CodigoProduto, " & _
        "EAN, " & _
        "NomeProduto, " & _
        "CFOP, " & _
        "NCM, " & _
        "CST, " & _
        "UnidadeComercial, " & _
        "ValorUnitarioComercializacao, " & _
        "ValorTotalBruto, " & _
        "tipodesconto, " & _
        "desconto, " & _
        "Valordesconto, " & _
        "QuantidadeComercial, " & _
        "pICMS, " & _
        "vBC, " & _
        "vICMS,  " & _
        "item, " & _
        "CodigoNota " & _
        " ) " & _
        "SELECT CodigoProduto, EAN, NomeProduto, CFOP, NCM, CST, UnidadeComercial, ValorUnitarioComercializacao, ValorTotalBruto, tipodesconto, desconto, Valordesconto, QuantidadeComercial, pICMS, vBC, vICMS, item, " & VarCodNota & " " & _
        "FROM NotaFiscalItens " & _
        "WHERE CodigoNota = " & varCodNotaOrigem & ";"
dbData.Execute sSQL

'preencher o grid dos itens com o pedido
Exibir_Itens


'finalizar o salvamento da nfe
LerDadosInserir
TbNotas.Update
'SomarProdutosNota

'PreencherGridNotas
    
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
frmNota.Enabled = True
frmDestinatario.Enabled = True
Tab_Totais.Enabled = True
Tab_Produtos.Enabled = True
frmItens.Enabled = True
cmdDuplicar.Enabled = False

Frm_NF.Tab = 0
Tab_Produtos.Tab = 0
Tab_Totais.Tab = 0
    
Call cmdRecalcular_Click

Exit Sub
Resume

'ErrLoad:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
'    Set TbNotas = Nothing

End Sub

Private Sub cmdDuplicarCFOP_Click()
'mudar cfop dos itens do grid itens
If cboNatureza.Text = "" Then Exit Sub
sSQL = "UPDATE NotaFiscalItens SET CFOP = '" & cboNatureza.Text & "', CST = '102' WHERE CodigoNota = " & Val(txtCodNota.Text)
dbData.Execute sSQL

Exibir_Itens
End Sub

Private Sub cmdEditar_Click()
'Clear_Controls
'LimparObjetosProduto
If vTipoEdicaoNFe = "Novo" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub
If vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub

If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub

If cmdEditar.Caption = "Editar" Then
    vTipoEdicaoNFe = "Edicao"
    RsOpen TbNotas, "SELECT *,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE (CASE WHEN Inutilizada = 1 THEN 'Inutilizada' ELSE 'Em Digitação' END) END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & GridNotas.TextMatrix(GridNotas.Row, 1)
    Load_Controls
    Frm_NF.Tab = 0
    cmdNovo.Enabled = False
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    frmNota.Enabled = True
    frmDestinatario.Enabled = True
    frmItens.Enabled = True
    Tab_Totais.Enabled = True
    Tab_Produtos.Enabled = True
    Exibir_Duplicatas
ElseIf cmdEditar.Caption = "Detalhar" Then
    vTipoEdicaoNFe = "Detalhar"
    RsOpen TbNotas, "SELECT *,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE (CASE WHEN Inutilizada = 1 THEN 'Inutilizada' ELSE 'Em Digitação' END) END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & GridNotas.TextMatrix(GridNotas.Row, 1)
    Load_Controls
    Frm_NF.Tab = 0
    cmdNovo.Enabled = True
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    frmNota.Enabled = False
    frmDestinatario.Enabled = False
    frmItens.Enabled = False
    Tab_Totais.Enabled = False
    Tab_Produtos.Enabled = False
    Exibir_Duplicatas
End If
Verificar_Duplicatas
End Sub


Private Sub cmdEnviarPDF_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
Dim vChave As String
Dim vDataProt As Date, dhDataProt As String
vChave = GridNotas.TextMatrix(GridNotas.Row, 8)
dhDataProt = GridNotas.TextMatrix(GridNotas.Row, 11)
vDataProt = Format(Left(dhDataProt, 10), "yyyy/mm/dd")

Dim NomeEmp As String, emailDestino As String, i As Integer, ComandoSQL As String

'parte de encontrar o arquivo
sSQL = "SELECT DiretorioXML, fantasia FROM Empresa"
Set rEmpresa = dbData.OpenRecordset(sSQL)

On Error GoTo deuErro
Dim sistNFe As snfe.Util
Set sistNFe = New snfe.Util

dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.path)
xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(vDataProt, "yyyymm") & "\" & vChave & "-procNFe.xml"
xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & vChave & ".pdf"

If Not Existe(xCaminhoXML) Then consultaNFe vChave, True

If Not Existe(xCaminhoXML) Then Exit Sub

'criar o arquivo
iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, xCancelada, True, "", True, False, False, False, True)

If Not Existe(xCaminhoPDF) Then MsgBox "Não existe o arquivo XML dessa venda nesse computador!", vbInformation, "Aviso do Sistema": Exit Sub

'envio do arquivo
emailDestino = InputBox("Informe o e-mail do destinatário", "Envio de Email", "")

If Not Vazio(emailDestino) Then
    'Call EnviaEmailPDF(emailDestino, xCaminhoXML)
    Call EnviaEmailPDF(emailDestino, xCaminhoPDF)
    DoEvents
End If

Exit Sub

deuErro:
    If InStr(1, Err.Description, "Exception") > 0 Then
       iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
       Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 1, xCancelada, True, "", True, False, False, False, True)
    Else
       MsgBox Err.Description, vbInformation
    End If
    Err.Clear
End Sub

Private Sub cmdEnviarXML_Click()
If GridNotas.rows <= 1 Then
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

IdNFProd = Val(GridNotas.TextMatrix(GridNotas.Row, 1))

sSQL = "SELECT ChaveDEAcesso, DataEmissao FROM NotaFiscal WHERE CodigoNota = " & IdNFProd
NFeChaveAcesso = SQLExecutaRetorno(sSQL, "ChaveDEAcesso", "")
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
corpoEmail = "Segue em anexo o arquivo PDF da NFe emitida. " & _
             "<br><br>" & _
             "Atenciosamente, " & _
             "<br><br>" & _
             "#nome_emitente#"
corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO"), SO_UM)

If (emailDest <> Empty) Then
   Screen.MousePointer = vbHourglass
   iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivo PDF referente a NFe emitida ", corpoEmail, pathAnexo, emailCC)
   Screen.MousePointer = vbDefault
End If

If iRetorno Then MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"

Set sistNFe = Nothing

Exit Sub
    Resume
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
corpoEmail = "Segue em anexo o arquivo XML da NFe emitida. " & _
             "<br><br>" & _
             "Atenciosamente, " & _
             "<br><br>" & _
             "#nome_emitente#"
corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RAZAO FROM empresa", "RAZAO"), SO_UM)

If (emailDest <> Empty) Then
   Screen.MousePointer = vbHourglass
   iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivo XML referente a NFe emitida ", corpoEmail, pathAnexo, emailCC)
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

Private Sub cmdEspelho_Click()
vPossuiErro = False

If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub

'verificar erros
If vPossuiErro = False Then VerificarDestinatarioEnviar Else Exit Sub
If vPossuiErro = False Then VerificarProdutosEnviar Else Exit Sub
If vPossuiErro = False Then CorrecoesBasicasNFe Else Exit Sub

Call CalcularTotalProdutos
Call CalcularDesconto
Call AtualizarValorICMS
Call CalcularICMSInterItens

If vPossuiErro = False Then
    vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))
    vSerieNota = (GridNotas.TextMatrix(GridNotas.Row, 14))
    
    DoEvents
    picAguarde.Visible = True
    iRetorno = TransmitirNFe(Val(vCodNota), Val(vSerieNota), False)
    If iRetorno Then
        
        On Error GoTo deuErro
          Dim sistNFe As snfe.Util
          Set sistNFe = New snfe.Util
     
        dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.path)
        xCaminhoXML = dirXML & "\nfe\arquivos\assinado\NFe" & NFeChaveAcesso & "-assinado.xml"
        xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
        
        If Not Existe(xCaminhoXML) Then
           picAguarde.Visible = False
           Exit Sub
        End If
        
        iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
        Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, False, False, "", True, False, False, False, True)
    End If
    picAguarde.Visible = False
    Call cmdExibirConNotas_Click
End If

    Exit Sub
deuErro:
    picAguarde.Visible = False
    If InStr(1, Err.Description, "Exception") > 0 Then
       iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
       Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 1, False, False, "", True, False, False, False, True)
    Else
       MsgBox Err.Description, vbInformation
    End If
    Err.Clear
End Sub

Private Sub cmdExibirConNotas_Click()
Dim totalRegistros As Long

'On Error GoTo ErrLoad

If cboFiltroNota.Text = "TODAS" Then
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal order by NumeroNota desc"
                sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    If cboConNotaCliente.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE NumeroNota = " & cboConNotaCliente.Text & " order by NumeroNota desc"
                sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE NumeroNota = " & cboConNotaCliente.Text & " order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "CLIENTE" Then
    If cboConNotaCliente.Text = "" Then Exit Sub
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " order by NumeroNota desc"
                sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " order by NumeroNota desc"
        'Debug.Print sSQL
ElseIf cboFiltroNota.Text = "DATAS" Then
    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) order by NumeroNota desc"
                sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND EmProcessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "MENSAL" Then

If cboConNotaMes.Text = "" Or cboConNotaAno.Text = "" Then Exit Sub

Dim vIndMes As Integer
Dim vWhere As String

If cboConNotaMes.ListCount = 0 Then
    If cboConNotaMes.Text = "janeiro" Then
        vIndMes = cboConNotaMes.ListIndex + 2
    ElseIf cboConNotaMes.Text = "fevereiro" Then
        vIndMes = cboConNotaMes.ListIndex + 3
    ElseIf cboConNotaMes.Text = "março" Then
        vIndMes = cboConNotaMes.ListIndex + 4
    ElseIf cboConNotaMes.Text = "abril" Then
        vIndMes = cboConNotaMes.ListIndex + 5
    ElseIf cboConNotaMes.Text = "maio" Then
        vIndMes = cboConNotaMes.ListIndex + 6
    ElseIf cboConNotaMes.Text = "junho" Then
        vIndMes = cboConNotaMes.ListIndex + 7
    ElseIf cboConNotaMes.Text = "julho" Then
        vIndMes = cboConNotaMes.ListIndex + 8
    ElseIf cboConNotaMes.Text = "agosto" Then
        vIndMes = cboConNotaMes.ListIndex + 9
    ElseIf cboConNotaMes.Text = "setembro" Then
        vIndMes = cboConNotaMes.ListIndex + 10
    ElseIf cboConNotaMes.Text = "outubro" Then
        vIndMes = cboConNotaMes.ListIndex + 11
    ElseIf cboConNotaMes.Text = "novembro" Then
        vIndMes = cboConNotaMes.ListIndex + 12
    ElseIf cboConNotaMes.Text = "dezembro" Then
        vIndMes = cboConNotaMes.ListIndex + 13
    End If
    
    vWhere = "(MONTH(DataEmissao) = " & vIndMes & ") "
Else
    vWhere = "(MONTH(DataEmissao) = " & cboConNotaMes.ListIndex + 1 & ") "
End If

    RsOpen TbConsulta, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF,  " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND Emprocessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by NumeroNota desc"
                sSQL = "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota,  FinalidadeEmissaoNFe, ChavedeAcesso, NumeroRecibo, DataHoraProcotolo, NumeroProtocolo, CodigoCorrentista, SerieNF, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND inutilizada = 0 AND Emprocessamento = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND inutilizada = 1 THEN 'Inutilizada' ELSE (CASE WHEN Enviada = 1 AND EmProcessamento = 1 THEN 'Em Processamento' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) END) END) AS Status, TipoCliente " & _
                    "FROM NotaFiscal WHERE " & vWhere & " AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by NumeroNota desc"
End If

printSQL = sSQL
'Debug.Print sSQL

'If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount

'lblQuantEnviada.Caption = Format(totalRegistros, "00")

LimparGridNotas
FormatarGridNotas TbConsulta

'lblTotalEnviada.Caption = Format(SomaGrid(GridNotas, 6), ocMONEY)
'lblTotalEnviadaCanceladas.Caption = Format(SomaGrid(GridNotas, 6), ocMONEY)

Dim soma As Currency
Dim contar As Integer
Dim i As Integer

'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Enviada" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantEnviada.Caption = Format(contar, "000")
lblTotalEnviada.Caption = Format(soma, ocMONEY)

'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Cancelada" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantCancelada.Caption = Format(contar, "000")
lblTotalCancelada.Caption = Format(soma, ocMONEY)

'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Inutilizada" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantInutilizada.Caption = Format(contar, "000")
lblTotalInutilizada.Caption = Format(soma, ocMONEY)
'Somar as vendas
soma = 0
contar = 0
With GridNotas
   For i = 1 To .rows - 1
      If .TextMatrix(i, 7) = "Em Digitação" Then
        'If .TextMatrix(i, 15) <> "SIM" Then
            contar = contar + 1
            soma = soma + CCur(.TextMatrix(i, 6))
        'End If
      End If
   Next
End With

lblQuantNaoEnviada.Caption = Format(contar, "000")
lblTotalNaoEnviada.Caption = Format(soma, ocMONEY)

cmdEditar.Caption = "Editar"
cmdEditar.Enabled = False
cmdTransmitir.Enabled = False
cmdCancelarNota.Enabled = False
cmdConsultar.Enabled = False
cmdInutilizar.Enabled = False
cmdImprimir.Enabled = False
cmdDuplicar.Enabled = False
'cmdReativar.Enabled = False
cmdCartaCorrecao.Enabled = False
cmdCopiarChave.Enabled = False
cmdEnviarXML.Enabled = False
cmdEnviarPDF.Enabled = False
cmdEspelho.Enabled = False

Exit Sub
Resume

'ErrLoad:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
'    Set TbConsulta = Nothing
End Sub

Private Sub cmdExibirPedidos_Click()
If cboIndicePedidos.Text = "" Then Exit Sub

Dim r As ADODB.Recordset
Dim totalRegistros As Long
   
If cboIndicePedidos.Text = "PEDIDO" Then
    If txtConCodPedido.Text = "" Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE pedidos.cod_pedido = " & txtConCodPedido & " AND (TIPO_PEDIDO <> 'ORÇAMENTO') ;"
ElseIf cboIndicePedidos.Text = "CLIENTE" Then
    If txtCodClientePedidos.Text = "" Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (cliente.codigo = " & txtCodClientePedidos.Text & ") AND (TIPO_PEDIDO <> 'ORÇAMENTO') ORDER BY pedidos.cod_pedido;"
ElseIf cboIndicePedidos.Text = "DATAS" Then
    If IsDate(mskInicialPedidos) = False Or IsDate(mskFinalPedidos) = False Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (pedidos.data_compra >= CONVERT(DATETIME, '" & Format(mskInicialPedidos.Text, ocDATA) & "', 103)) AND (pedidos.data_compra <= CONVERT(DATETIME, '" & Format(mskFinalPedidos.Text, ocDATA) & "', 103)) AND (TIPO_PEDIDO <> 'ORÇAMENTO') ORDER BY pedidos.cod_pedido;"
ElseIf cboIndicePedidos.Text = "MENSAL" Then
    If cboMesPedidos.Text = "" Or cboAnoPedidos.Text = "" Then Exit Sub 'var_tipoPGTO
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (MONTH(pedidos.data_compra) = " & cboMesPedidos.ListIndex + 1 & ") AND (YEAR(pedidos.data_compra) = " & cboAnoPedidos & ") ORDER BY pedidos.cod_pedido;"
Else
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE pedidos.cod_pedido = '0';"
End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGridPedidos r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'MOSTRAR A QUANTIDADE REGISTROS
   lblQuantPedidos.Caption = Format(totalRegistros, "00")

End Sub

Private Sub cmdFecharCCe_Click()
frmCorreção.Visible = False
End Sub



Private Sub cmdImprimirConsulta_Click()
'colocar o nome da maquina na barra de status
Dim oIni As Ini
Dim var_Impressora As String
'Dim r As ADODB.Recordset

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_NFe_Consulta.Relatorio.Recordset = r
REL_NFe_Consulta.dfQuant.Caption = lblQuantEnviada.Caption
REL_NFe_Consulta.dfBruto.Caption = lblTotalEnviada.Caption

If cboFiltroNota.Text = "MENSAL" Then
   REL_NFe_Consulta.dfTipo.Caption = "Tipo: Mês = " & cboConNotaMes.Text & "/" & cboConNotaAno.Text
ElseIf cboFiltroNota.Text = "DATAS" Then
   REL_NFe_Consulta.dfTipo.Caption = "Tipo: Datas = " & mskConNotaInicial.Text & " à " & mskConNotaFinal.Text
ElseIf cboFiltroNota.Text = "CLIENTE" Then
   REL_NFe_Consulta.dfTipo.Caption = "Tipo: Cliente = " & cboConNotaCliente.Text & ""
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
   REL_NFe_Consulta.dfTipo.Caption = "Tipo: Nota Fiscal Nº " & cboConNotaCliente.Text & ""
Else
   REL_NFe_Consulta.dfTipo.Caption = "Tipo: Todas as notas"
End If

REL_NFe_Consulta.Relatorio.Ativar
Unload REL_NFe_Consulta
Me.Show 1
End Sub

Private Sub cmdInutilizar_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub

Dim codPedido As String, nNota As String, CNPJ As String, idInutilizacao As Long
Dim sSQL As String, IdNFProd As Long

vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))
codPedido = (GridNotas.TextMatrix(GridNotas.Row, 1))

dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM Empresa", "DiretorioXML", App.path)
dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
CNPJ = SQLExecutaRetorno("SELECT CNPJ FROM Empresa", "CNPJ", "")

sSQL = "SELECT CodigoNota FROM NotaFiscal WHERE CodigoNota  = " & codPedido
IdNFProd = SQLExecutaRetorno(sSQL, "CodigoNota", 0)
If IdNFProd > 0 Then
   sSQL = "SELECT CodigoInutilizacao FROM NFeInutilizacao WHERE NumeroNotaInicial = " & IdNFProd
   idInutilizacao = SQLExecutaRetorno(sSQL, "CodigoInutilizacao", 0)
   sSQL = "SELECT NumeroNota FROM NotaFiscal WHERE CodigoNota = " & IdNFProd
   nNota = SQLExecutaRetorno(sSQL, "NumeroNota", "0")
   If idInutilizacao = 0 Then
      sSQL = "SELECT ISNULL(MAX(CodigoInutilizacao), 0) r FROM NFeInutilizacao"
      idInutilizacao = SQLExecutaRetorno(sSQL, "r", 0) + 1
      sSQL = "INSERT INTO [NFeInutilizacao] ([CodigoInutilizacao],[Ano],[NumeroNotaInicial],[NumeroNotaFinal],[Justificativa]) Values " & _
             "(" & idInutilizacao & ", " & Format(Date, "yy") & ", " & nNota & ", " & nNota & ", 'ERRO AO TRANSMITIR NOTA, PERDA DE SEQUENCIA')"
      vgDb.Execute sSQL
   End If
   Dim sistNFe As snfe.Util
   Set sistNFe = New snfe.Util
   iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
   iRetorno = sistNFe.InutilizarNumeracao(Format(Date, "yyyy"), CNPJ, "ERRO AO TRANSMITIR NOTA, PERDA DE SEQUENCIA", nNota, nNota, 1, xCaminhoXML)
   cStat = sistNFe.retInutilizacao.infInut.cStat
   NFeMotivo = sistNFe.retInutilizacao.infInut.xMotivo
   NFeDataHora = sistNFe.retInutilizacao.infInut.dhRecbto
   NFeNumeroProtocolo = sistNFe.retInutilizacao.infInut.nProt
   If cStat = 102 Then
      sSQL = "UPDATE NFeInutilizacao SET " & _
             "Enviada = 1, " & _
             "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
             "DataHora = '" & NFeDataHora & "' " & _
             "WHERE CodigoInutilizacao = " & idInutilizacao
      vgDb.Execute sSQL
      sSQL = "UPDATE NotaFiscal SET Enviada = 1, Inutilizada = 1 WHERE CodigoNota = " & IdNFProd
      vgDb.Execute sSQL
      MsgBox CStr(cStat) & " - " & NFeMotivo, vbInformation + vbOKOnly, "INUTILIZAÇÃO"
   Else
      MsgBox CStr(cStat) & " - " & NFeMotivo, vbCritical + vbOKOnly, "ERRO - INUTILIZAÇÃO"
   End If
   
   Set sistNFe = Nothing
End If
    Call cmdExibirConNotas_Click
End Sub

Private Sub cmdNovo_Click()
'On Error GoTo ErrLoad
vTipoEdicaoNFe = "Novo"

'pegando o numero correto da nota
'Dim var_NumeroNota As Integer
Dim ConsultaSQL As String
Dim tbNota As ADODB.Recordset

      ConsultaSQL = "SELECT ISNULL(MAX(numeronota), 0) AS Maior_nota FROM NotaFiscal"
      Set tbNota = dbData.OpenRecordset(ConsultaSQL)
      'If Not tbNota.BOF Then var_NumeroNota = tbNota("ultima_nota") + 1
      
'preecher objetos do form
Dim totalRegistros As Long
RsOpen TbNotas, "SELECT *,  " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM NotaFiscal"
                
If TbNotas.RecordCount > 0 Then totalRegistros = TbNotas.RecordCount

    LimparObjetosNota
    LimparObjetosDestinatario
    LimparObjetosProduto
    LimparObjetosNotaTotais
    LimparObjestosNotaOutros
    LimparGridItensNota
    'LimparGridItensNota
    
    If TbNotas.EOF And TbNotas.BOF Then
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = "1"
        txtSerie.Text = "1"
    Else
        TbNotas.MoveLast
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = TbNotas("CodigoNota") + 1
        txtSerie.Text = TbNotas("serienf")
    End If
    
    ' Cria registro em branco para que os itens possam ser adicionados antes do Salvar
    RsOpen TbNotas, "SELECT * FROM NotaFiscal WHERE 1 = 0"
    TbNotas.AddNew
    TbNotas("CodigoNota") = Val(txtCodNota.Text)
    TbNotas("NumeroNota") = Val(txtNumNota.Text)
    TbNotas("SerieNF") = 1
    TbNotas("DataEmissao") = Now()
    TbNotas("DataSaida") = Now()
    TbNotas("HoraSaida") = Now()
    TbNotas.Update
    
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormaPgto.Text = "01 = Dinheiro"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
    'cboModFrete.Text = ""
    
    'txtValorFrete.Text = "0,00"
    'txtValorOutrasDespesas.Text = "0,00"
    'txtVolPesoBruto.Text = "0,00"
    'txtVolPesoLiquido.Text = "0,00"
    'txtValorSeguro.Text = "0,00"
    'txtBaseICMSST.Text = "0,00"
    'txtBaseICMS.Text = "0,00"
    'txtValorICMS.Text = "0,00"
    'txtBaseICMSST.Text = "0,00"
    'txtValorICMSST.Text = "0,00"
    'txtValorIPI.Text = "0,00"
    'txtValorDesconto.Text = "0,00"
    
    cboModFrete.Text = "9 - SEM FRETE"
    cboTipoNota.Text = "1 - SAÍDA"
    cboFinalidade.Text = "1 - NFe NORMAL"
    cboConsumidorFinal.Text = "1 - SIM"
    mskEmissao = Format(Date, "dd/mm/yyyy")
    mskSaida = Format(Date, "dd/mm/yyyy")
    mskHora = Format(Time(), "HH:MM:ss")
    
    
    'LimparObjetosProduto
    
    TbNotas.AddNew
    cmdNovo.Enabled = False
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    frmNota.Enabled = True
    frmDestinatario.Enabled = True
    frmItens.Enabled = True
    Tab_Totais.Enabled = True
    Tab_Produtos.Enabled = True

    If vTipoCRT = 3 Then
        txtInfComple.Text = ""
    Else
        txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
    End If
    
    Mostrar_AliqUF
    cboTipoDest.SetFocus

Exit Sub
Resume

'ErrLoad:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
'    Set TbNotas = Nothing
End Sub
Private Sub Calcular_Parcelas()
If txtTotalDup.Text = "0,00" Or txtTotalDup.Text = "" Then Exit Sub
If txtNumParcDup.Text = "" Then txtNumParcDup.Text = "1"
'If txtNumParcDup.Text = "0" Or txtNumParcDup.Text = "" Then Exit Sub

Dim vValorTotal As Currency
Dim vQuant As Integer
Dim vResultado As Currency

vValorTotal = txtTotalDup.Text
vQuant = txtNumParcDup.Text

vResultado = CCur(vValorTotal / vQuant)
txtValorParcDup = Format(vResultado, ocMONEY)
End Sub




Private Sub cmdRecalcular_Click()
'Call MostrarValorItens  'desativei pq estava calculando o valortotalbruto errado
'Call AtualizarValorICMS
'If Left(cboDestOperacao.Text, 1) = "2" Then Call CalcularICMSInterItens
'Call DistribuirFrete
'Call DistribuirSeguro
'Call DistribuirOutros
'Call DistribuirDesconto
'Call AtualizarTotaisNota
End Sub

Private Sub cmdRemoverDuplicatas_Click()
If ShowMsg("Deseja realmente excluir todas as duplicatas dessa Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
dbData.Execute "DELETE FROM NotaFiscalParcelas WHERE CodigoNota = " & Val(txtCodNota.Text)
Exibir_Duplicatas
End Sub

Private Sub cmdRemoverItem_Click()
Dim vTotal As Double

'On Error GoTo erro

    'sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    'RsOpen Tb, sSQL


    'vgDb.BeginTrans
    
    'Tb.AddNew 'insere os dados
    'Load_Data_Itens
    'Tb.Update
    
    'vgDb.CommitTrans
    
    'Limpa_Tudo Me ' limpa tudo
    If ShowMsg("Deseja remover o item: " & GridNotasItens.TextMatrix(GridNotasItens.Row, 4) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    
    dbData.Execute "DELETE FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND (CodigoProduto = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 3) & ") AND (ITEM = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 1) & ");"

    'Call DistribuirFrete
    'Call DistribuirOutros
    'Call DistribuirSeguro
    'Call CalcularIPI
    'Call MostrarValorProdutos
    'Call CalcularDesconto
    'Call AtualizarValorICMS
    'If Left(cboDestOperacao.Text, 1) = "2" Then Call CalcularICMSInterItens
    'Call MostrarValorBaseICMS
    'Call MostrarValorNota

    
   ' sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    'vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    'sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)

    'SQLExecuta sSQL
    
    Exibir_Itens
    KeyCode = 0
    TipoSelecaoConsulta = "0"
    AtualizarTotaisNota
    cboDescricao.SetFocus
Exit Sub
End Sub


Private Sub cmdSalvar_Click()
flag = False

'On Error GoTo Err_Grava
If txtCodCliente.Text = "1" Then MsgBox "NÃO É PERMITIDO FAZER NFE PARA ESSE CLIENTE." & Chr(13) & "Selecione outro cliente!", vbInformation, "Aviso do Sistema": cboCliente.SetFocus: Exit Sub
If mskHora.Text = "" Then MsgBox "O campo hora é obrigatório!", vbInformation, "Aviso do Sistema": mskHora.SetFocus: Exit Sub
If Not IsDate(mskEmissao) Then MsgBox "O campo hora é obrigatório!", vbInformation, "Aviso do Sistema": mskEmissao.SetFocus: Exit Sub
If Not IsDate(mskSaida) Then MsgBox "O campo hora é obrigatório!", vbInformation, "Aviso do Sistema": mskSaida.SetFocus: Exit Sub
If txtCodCliente.Text = "" Then MsgBox "O campo CLIENTE é obrigatório.", vbCritical, "Online Commerce": cboCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo Destino é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
'If txtCodObservacao.Text = "" Then MsgBox "O campo mensagem é obrigatório.", vbCritical, "Online Commerce": txtCodObservacao.SetFocus: Exit Sub
'If txtTotaldosProdutos.Text = "" Then cmdRecalcular_Click
cmdRecalcular_Click

'VerificarDestinatario

If vTipoEdicaoNFe = "Novo" Then

    resp = MsgBox("Confirma inclusão ?", 36, Titulo)
    flag = True
    If resp <> 6 Then Exit Sub
    
    RsOpen TbNotas, "SELECT * FROM NotaFiscal WHERE CodigoNota = " & Val(txtCodNota.Text)
    Load_Data
    TbNotas.Update
    vgDb.CommitTrans

ElseIf vTipoEdicaoNFe = "Edicao" Then
        
    If TbNotas.EditMode = 2 Then
       resp = MsgBox("Confirma inclusão ?", 36, Titulo)
       flag = True
       If resp <> 6 Then Exit Sub
    Else
       resp = MsgBox("Confirma alteração ?", 36, Titulo)
       flag = False
       If resp <> 6 Then Exit Sub
    End If
    
    'If txtTotaldosProdutos.Text = "" Then
    '    cmdRecalcular_Click
    'End If
    
    Load_Data
    TbNotas.Update
    
    vgDb.CommitTrans

End If

    cmdNovo.Enabled = True
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    frmNota.Enabled = False
    frmDestinatario.Enabled = False
    Tab_Totais.Enabled = False
    Tab_Produtos.Enabled = False
    frmItens.Enabled = False
    Tab_Totais.Tab = 0
    Tab_Produtos.Tab = 0
    
    'Clear_Controls
    LimparObjetosNota
    LimparObjetosDestinatario
    LimparObjetosProduto
    LimparObjetosNotaTotais
    LimparObjestosNotaOutros
    LimparGridItensNota
    vTipoEdicaoNFe = ""
    Call cmdExibirConNotas_Click

Exit Sub
Resume
'Err_Grava:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
End Sub
Private Sub cmdImprimir_Click()
If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
Dim vChave As String
Dim vDataProt As Date, dhDataProt As String
vChave = GridNotas.TextMatrix(GridNotas.Row, 8)
dhDataProt = GridNotas.TextMatrix(GridNotas.Row, 11)
vDataProt = Format(Left(dhDataProt, 10), "yyyy/mm/dd")

   On Error GoTo deuErro
     Dim sistNFe As snfe.Util
     Set sistNFe = New snfe.Util
     
     dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.path)
     xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(vDataProt, "yyyymm") & "\" & vChave & "-procNFe.xml"
     xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & vChave & ".pdf"
     
     If Not Existe(xCaminhoXML) Then consultaNFe vChave, True
     
     If Not Existe(xCaminhoXML) Then Exit Sub
     
     iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
     Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, xCancelada, False, "", True, False, False, False, True)
     
     Exit Sub
deuErro:
    If InStr(1, Err.Description, "Exception") > 0 Then
       iRetorno = ConfiguraDLLNFeNFCe(55, "1", sistNFe)
       Call sistNFe.DANFeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 1, xCancelada, False, "", True, False, False, False, True)
    Else
       MsgBox Err.Description, vbInformation
    End If
    Err.Clear
End Sub

Private Sub cmdTransmitir_Click()
vPossuiErro = False

If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub
If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub

'verificar erros
If vPossuiErro = False Then VerificarDestinatarioEnviar Else Exit Sub
If vPossuiErro = False Then VerificarProdutosEnviar Else Exit Sub
If vPossuiErro = False Then CorrecoesBasicasNFe Else Exit Sub

Call CalcularTotalProdutos
Call CalcularDesconto
Call AtualizarValorICMS
Call CalcularICMSInterItens

If vPossuiErro = False Then
    vCodNota = (GridNotas.TextMatrix(GridNotas.Row, 1))
    vSerieNota = (GridNotas.TextMatrix(GridNotas.Row, 14))
    
    DoEvents
    picAguarde.Visible = True
    iRetorno = TransmitirNFe(Val(vCodNota), Val(vSerieNota), True)
    If iRetorno Then
       SQL = "UPDATE NotaFiscal SET " & _
             "Enviada = 1 " & _
             "WHERE CodigoNota = " & Val(vCodNota)
       vgDb.Execute SQL
       'RsOpen TbNotas, "SELECT *,  " & _
                       "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                       "FROM NotaFiscal WHERE CodigoNota = " & Val(vCodNota)
       'Load_Controls
       'FormatarGridNotas TbNotas
    End If
    picAguarde.Visible = False
Call cmdExibirConNotas_Click
End If
End Sub
Private Sub FormatarGridPedidos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridPedidos
       .Clear
       .Cols = 7
       .rows = 2
           
       .ColWidth(0) = 0
       .ColWidth(1) = 900
       .ColWidth(2) = 900
       .ColWidth(3) = 1100
       .ColWidth(4) = 4000
       .ColWidth(5) = 2000
       .ColWidth(6) = 2000
 
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "NFE"
      .TextMatrix(0, 3) = "COMPRA"
      .TextMatrix(0, 4) = "CLIENTE"
      .TextMatrix(0, 5) = "TIPO"
      .TextMatrix(0, 6) = "VALOR"

      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("var_codped")
            .TextMatrix(.rows - 1, 2) = rTabela("status")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 4) = rTabela("var_Nome")
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_tipoPGTO"))
            .TextMatrix(.rows - 1, 6) = Format(rTabela("var_total"), ocMONEY)
            
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
   
'   lblValor.Caption = Format(SomaGrid(GridPedidos, 5), ocMONEY)
End Sub





Private Sub Form_Activate()
Exibir_Cliente
Exibir_Itens
End Sub



Private Sub Grid_Correcao_Click()
i = Grid_Correcao.Row
If Grid_Correcao.TextMatrix(i, 8) = "ENVIADO" Then
    cmdCCeImprimir.Enabled = True
    cmdCCeTransmitir.Enabled = False
    cmdCCeExcluir.Enabled = False
ElseIf Grid_Correcao.TextMatrix(i, 8) = "NÃO ENVIADO" Then
    cmdCCeImprimir.Enabled = False
    cmdCCeTransmitir.Enabled = True
    cmdCCeExcluir.Enabled = True
End If
End Sub

Private Sub GridNotas_Click()
i = GridNotas.Row
If GridNotas.TextMatrix(i, 7) = "Enviada" Then
    cmdEditar.Caption = "Detalhar"
    cmdEditar.Enabled = True
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = True
    cmdConsultar.Enabled = True
    cmdInutilizar.Enabled = False
    cmdImprimir.Enabled = True
    cmdDuplicar.Enabled = True
    'cmdReativar.Enabled = True
    cmdCartaCorrecao.Enabled = True
    cmdCopiarChave.Enabled = True
    cmdEnviarXML.Enabled = True
    cmdEnviarPDF.Enabled = True
    cmdEspelho.Enabled = False
ElseIf GridNotas.TextMatrix(i, 7) = "Em Digitação" Then
    cmdEditar.Enabled = True
    cmdEditar.Caption = "Editar"
    cmdTransmitir.Enabled = True
    cmdCancelarNota.Enabled = False
    cmdConsultar.Enabled = False
    cmdInutilizar.Enabled = True
    cmdImprimir.Enabled = False
    cmdDuplicar.Enabled = False
    'cmdReativar.Enabled = False
    cmdCartaCorrecao.Enabled = False
    cmdCopiarChave.Enabled = False
    cmdEnviarXML.Enabled = False
    cmdEnviarPDF.Enabled = False
    cmdEspelho.Enabled = True
    If Len(GridNotas.TextMatrix(i, 8)) = 44 Then cmdConsultar.Enabled = True
ElseIf GridNotas.TextMatrix(i, 7) = "Cancelada" Then
    cmdEditar.Enabled = True
    cmdEditar.Caption = "Detalhar"
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdConsultar.Enabled = True
    cmdInutilizar.Enabled = False
    cmdImprimir.Enabled = False
    cmdDuplicar.Enabled = True
    'cmdReativar.Enabled = False
    cmdCartaCorrecao.Enabled = False
    cmdCopiarChave.Enabled = True
    cmdEnviarXML.Enabled = True
    cmdEnviarPDF.Enabled = True
    cmdEspelho.Enabled = False
ElseIf GridNotas.TextMatrix(i, 7) = "Inutilizada" Then
    cmdEditar.Enabled = True
    cmdEditar.Caption = "Detalhar"
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdConsultar.Enabled = False
    cmdInutilizar.Enabled = False
    cmdImprimir.Enabled = False
    cmdDuplicar.Enabled = True
    'cmdReativar.Enabled = False
    cmdCartaCorrecao.Enabled = False
    cmdCopiarChave.Enabled = False
    cmdEnviarXML.Enabled = False
    cmdEnviarPDF.Enabled = False
    cmdEspelho.Enabled = False
ElseIf GridNotas.TextMatrix(i, 7) = "Em Processamento" Then
    cmdEditar.Enabled = False
    cmdEditar.Caption = "Detalhar"
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdConsultar.Enabled = True
    cmdInutilizar.Enabled = False
    cmdImprimir.Enabled = False
    cmdDuplicar.Enabled = False
    'cmdReativar.Enabled = False
    cmdCartaCorrecao.Enabled = False
    cmdCopiarChave.Enabled = True
    cmdEnviarXML.Enabled = False
    cmdEnviarPDF.Enabled = False
    cmdEspelho.Enabled = False
End If
End Sub





Private Sub mskConNotaFinal_GotFocus()
SelectControl mskConNotaFinal
End Sub

Private Sub mskConNotaFinal_KeyPress(KeyAscii As Integer)
mskConNotaFinal.Mask = "##/##/##"
End Sub


Private Sub mskConNotaFinal_LostFocus()
If mskConNotaFinal.Text = "" Or mskConNotaFinal.Text = "__/__/__" Then
   mskConNotaFinal.Mask = ""
   mskConNotaFinal.Text = ""
   Exit Sub
Else
   If IsDate(mskConNotaFinal.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskConNotaFinal.SetFocus
      SelectControl mskConNotaFinal
   End If
End If
End Sub


Private Sub mskConNotaInicial_GotFocus()
SelectControl mskConNotaInicial
End Sub


Private Sub mskConNotaInicial_KeyPress(KeyAscii As Integer)
mskConNotaInicial.Mask = "##/##/##"
End Sub


Private Sub mskConNotaInicial_LostFocus()
If mskConNotaInicial.Text = "" Or mskConNotaInicial.Text = "__/__/__" Then
   mskConNotaInicial.Mask = ""
   mskConNotaInicial.Text = ""
   Exit Sub
Else
   If IsDate(mskConNotaInicial.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskConNotaInicial.SetFocus
      SelectControl mskConNotaInicial
   End If
End If

End Sub

Private Sub mskEmissao_KeyPress(KeyAscii As Integer)
mskEmissao.Mask = "##/##/####"
End Sub

Private Sub mskHora_KeyPress(KeyAscii As Integer)
mskHora.Mask = "##:##:##"
End Sub

Private Sub mskInicioDup_GotFocus()
SelectControl mskInicioDup
End Sub


Private Sub mskInicioDup_KeyPress(KeyAscii As Integer)
If Not IsDate(mskInicioDup.Text) Then Exit Sub
mskInicioDup.Mask = "##/##/##"
End Sub


Private Sub mskSaida_KeyPress(KeyAscii As Integer)
mskSaida.Mask = "##/##/####"
End Sub

Private Sub Tab_Totais_Click(PreviousTab As Integer)
If Tab_Totais.Tab = 0 Then
    'txtNome.SetFocus
ElseIf Tab_Totais.Tab = 1 Then
    'mskAdmissao.SetFocus
ElseIf Tab_Totais.Tab = 2 Then
    'Tab_Totais.Tab = 0
ElseIf Tab_Totais.Tab = 3 Then
    Tab_transp.Tab = 0
End If
End Sub




Private Sub TxtCodCliente_Change()
Dim TbClientes As New ADODB.Recordset
Dim TbEmpresa As New ADODB.Recordset

If txtCodCliente.Text = "" Then
    'txtCliEndereco.Text = ""
    'txtCliNum.Text = ""
    'txtCliBairro.Text = ""
    'txtCliCidade.Text = ""
    'txtCliUF.Text = ""
    'txtCliIBGE.Text = ""
    'txtCliCPF.Text = ""
    'txtCliIE.Text = ""
    'Exit Sub
Else
    If cboTipoDest.Text = "FORNECEDOR" Then
        RsOpen TbClientes, "SELECT codigo, TipoContribuinte, estado FROM fornecedor WHERE codigo = " & Val(txtCodCliente.Text)
    Else
        RsOpen TbClientes, "SELECT codigo, TipoContribuinte, estado FROM cliente WHERE codigo = " & Val(txtCodCliente.Text)
    End If
    
    
    If Not TbClientes.EOF And Not TbClientes.BOF Then
        If TbClientes("TipoContribuinte") = 1 Then
            cboTipoContribuinte.Text = "1 - CONTRIBUINTE ICMS"
        ElseIf TbClientes("TipoContribuinte") = 2 Then
            cboTipoContribuinte.Text = "2 - CONTRIBUINTE ISENTO"
        ElseIf TbClientes("TipoContribuinte") = 9 Then
            cboTipoContribuinte.Text = "9 - NÃO CONTRIBUINTE"
        End If
        
        vUFDest = ValidateNull(TbClientes("estado"))
        
        sSQL = "SELECT ESTADO FROM empresa"
        Set TbEmpresa = dbData.OpenRecordset(sSQL)
        
        If Not TbEmpresa.EOF Then
            If TbClientes("estado") = TbEmpresa("estado") Then
                cboDestOperacao.Text = "1 - Operação Interna"
            ElseIf TbClientes("estado") <> TbEmpresa("estado") Then
                cboDestOperacao.Text = "2 - Operação Interestadual"
            End If
        End If
        'txtCliEndereco.Text = ValidateNull(TbClientes("endereco"))
        'txtCliNum.Text = ValidateNull(TbClientes("numero"))
        'txtCliBairro.Text = ValidateNull(TbClientes("bairro"))
        'txtCliCidade.Text = ValidateNull(TbClientes("cidade"))
        'txtCliUF.Text = ValidateNull(TbClientes("estado"))
        
        'If txtCliUF.Text <> "" Then
        '    If txtCliUF.Text = "MA" Then
        '        vAliqUFDest = Format(18, "#0.00")
        '    ElseIf txtCliUF.Text = "BA" Then
        '        vAliqUFDest = Format(18, "#0.00")
        '    ElseIf txtCliUF.Text = "SP" Then
        '        vAliqUFDest = Format(18, "#0.00")
        '    Else
        '        vAliqUFDest = Format(18, "#0.00")
        '    End If
        'End If
        'txtCliIBGE.Text = ValidateNull(TbClientes("CodigoIBGE"))
        'txtCliCPF.Text = ValidateNull(TbClientes("cpf"))
        'txtCliIE.Text = ValidateNull(TbClientes("ie"))
    End If
End If
End Sub

Private Sub txtCodTransporte_Change()
If txtCodTransporte.Text = "" Then Exit Sub

Dim TbTransportadora As New ADODB.Recordset

'On Error GoTo erro
   
RsOpen TbTransportadora, "select * from transportadora where codigo=" & Val(txtCodTransporte.Text)

If Not TbTransportadora.EOF Then
    txtCodTransporte.Text = TbTransportadora("codigo")
    cboTransporte.Text = TbTransportadora("razao")
    vTranspCNPJ = TbTransportadora("CNPJ")
    vTranspEnd = TbTransportadora("Endereco")
    vTranspCidade = TbTransportadora("Cidade")
    vTranspUF = TbTransportadora("estado")
    vTranspIE = TbTransportadora("ie")
Else
    'txtCodTransporte.Text = ""
    cboTransporte.Text = ""
    vTranspCNPJ = ""
    vTranspEnd = ""
    vTranspCidade = ""
    vTranspUF = ""
    vTranspIE = ""
End If

'erro:
'MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
Calcular_TotalItem
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
'Exit Sub
If KeyCode = 38 Then
   If GridNotasItens.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
   GridNotasItens.Row = iRow - 1
   GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
   GridNotasItens_Click

ElseIf KeyCode = 40 Then
   If GridNotasItens.rows = GridNotasItens.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
   GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
   GridNotasItens.Row = iRow + 1
   GridNotasItens_Click
End If
End Sub
Private Sub txtEdit_LostFocus()
Dim sVal      As String
Dim sItem     As String
Dim sCodProd  As String
Dim curVBC    As Currency
Dim curVICMS  As Currency
Dim curVBCST  As Currency
Dim curVICMSST As Currency
Dim curVIPI   As Currency
Dim curSubTot As Currency
Dim dblPICMS  As Double
Dim dblPICMSST As Double
Dim dblPRedBC As Double
Dim dblPIPI   As Double
Dim dblMVA    As Double

txtEdit.Visible = False
sVal     = Trim(txtEdit.Text)
sItem    = GridNotasItens.TextMatrix(iRow, 1)
sCodProd = GridNotasItens.TextMatrix(iRow, 3)

If sItem = "" Then Exit Sub

Select Case iCol

    Case 2 ' EAN
        sVal = Replace(sVal, " ", "")
        If sVal = "" Or UCase(sVal) = "SEM GTIN" Then
            sVal = "SEM GTIN"
        Else
            If Not IsNumeric(sVal) Then
                MsgBox "EAN deve conter apenas dígitos!", vbInformation, "Aviso"
                Exit Sub
            End If
            If Len(sVal) <> 8 And Len(sVal) <> 13 Then
                MsgBox "EAN deve ter 8 ou 13 dígitos!", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET EAN = '" & sVal & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        dbData.Execute "UPDATE Produtos SET EAN = '" & sVal & "', COD_BARRA = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 5 ' UND
        If sVal = "" Then
            MsgBox "Unidade não pode ser vazia!", vbInformation, "Aviso"
            Exit Sub
        End If
        sVal = UCase(sVal)
        Dim sListaUND As String
        sListaUND = "|UN|PC|KG|CX|PA|PT|LT|ML|GR|MG|DZ|FD|RL|JG|KT|LA|GL|BD|SC|PR|M2|M3|CT|EX|BJ|DI|MET|"
        If InStr(sListaUND, "|" & sVal & "|") = 0 Then
            MsgBox "Unidade '" & sVal & "' inválida!" & vbCrLf & "Aceitas: UN PC KG CX PA PT LT ML GR DZ FD RL JG KT LA GL BD SC PR M2 M3 CT EX BJ DI MET", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET UnidadeComercial = '" & sVal & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        dbData.Execute "UPDATE Produtos SET unid_medida = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal
        ' Reformatar QTDE conforme unidade
        Dim sQtdAtual As String
        sQtdAtual = GridNotasItens.TextMatrix(iRow, 10)
        If sVal = "KG" Or sVal = "GR" Or sVal = "MG" Then
            GridNotasItens.TextMatrix(iRow, 10) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), ocPESO)
        Else
            GridNotasItens.TextMatrix(iRow, 10) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), "###,###,##0")
        End If

    Case 6 ' NCM
        sVal = Replace(sVal, ".", "")
        If sVal <> "" Then
            If Len(sVal) <> 8 Or Not IsNumeric(sVal) Then
                MsgBox "NCM deve ter 8 dígitos!", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET NCM = '" & sVal & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        dbData.Execute "UPDATE Produtos SET NCM = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 7 ' CFOP
        If sVal = "" Or Len(sVal) <> 4 Or Not IsNumeric(sVal) Then
            MsgBox "CFOP deve ter 4 dígitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET CFOP = " & Val(sVal) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 8 ' CST
        If sVal = "" Or Len(sVal) <> 3 Then
            MsgBox "CST deve ter 3 dígitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET CST = '" & sVal & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 17 ' %ICMS
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Alíquota ICMS inválida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPICMS = Val(sVal)
        curVBC   = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 16), ".", ""), ",", ".")))
        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET pICMS = " & FSQL(dblPICMS, 4) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 17) = FormatNumber(dblPICMS, 2)
        GridNotasItens.TextMatrix(iRow, 18) = FormatNumber(curVICMS, 2)

    Case 19 ' %RED BC
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Redução BC inválida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPRedBC = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 15), ".", ""), ",", ".")))
        curVBC    = CCur(Format(curSubTot * (1 - dblPRedBC / 100), "0.00"))
        dblPICMS  = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 17), ".", ""), ",", "."))
        curVICMS  = CCur(Format(curVBC * dblPICMS / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = " & FSQL(dblPRedBC, 4) & ", vBC = " & FSQL(curVBC, 2) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 19) = FormatNumber(dblPRedBC, 2)
        GridNotasItens.TextMatrix(iRow, 16) = FormatNumber(curVBC, 2)
        GridNotasItens.TextMatrix(iRow, 18) = FormatNumber(curVICMS, 2)

    Case 21 ' %ICMSST
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Alíquota ICMS-ST inválida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPICMSST  = Val(sVal)
        curVBCST    = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))
        curVICMS    = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 18), ".", ""), ",", ".")))
        curVICMSST  = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
        If curVICMSST < 0 Then curVICMSST = 0
        dbData.Execute "UPDATE NotaFiscalItens SET pICMSST = " & FSQL(dblPICMSST, 4) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(dblPICMSST, 2)
        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMSST, 2)

    Case 23 ' MVA ST
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then
            MsgBox "MVA inválido (deve ser >= 0)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblMVA      = Val(sVal)
        curSubTot   = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 15), ".", ""), ",", ".")))
        curVIPI     = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 26), ".", ""), ",", ".")))
        curVBCST    = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), "0.00"))
        dblPICMSST  = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 21), ".", ""), ",", "."))
        curVICMS    = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 18), ".", ""), ",", ".")))
        curVICMSST  = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
        If curVICMSST < 0 Then curVICMSST = 0
        dbData.Execute "UPDATE NotaFiscalItens SET pMVAST = " & FSQL(dblMVA, 4) & ", vBCST = " & FSQL(curVBCST, 2) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(dblMVA, 2)
        GridNotasItens.TextMatrix(iRow, 20) = FormatNumber(curVBCST, 2)
        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMSST, 2)

    Case 24 ' CST IPI
        If sVal = "" Or Len(sVal) <> 2 Or Not IsNumeric(sVal) Then
            MsgBox "CST IPI deve ter 2 dígitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET IPICST = '" & sVal & "', IPIcEnq = '999' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 25 ' %IPI
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Alíquota IPI inválida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPIPI   = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 15), ".", ""), ",", ".")))
        curVIPI   = CCur(Format(curSubTot * dblPIPI / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET IPIpIPI = " & FSQL(dblPIPI, 4) & ", IPIvIPI = " & FSQL(curVIPI, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(dblPIPI, 2)
        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVIPI, 2)

End Select

AtualizarTotaisNota
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then SendK vbKeyTab
If KeyCode = vbKeyUp Then
    SendK vbKeyTab
    KeyCode = 0
ElseIf KeyCode = vbKeyDown Then
    SendK vbKeyTab
    KeyCode = 0
End If

If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
Set moCombo = New cComboHelper
Frm_NF.Tab = 0
Tab_Produtos.Tab = 0
Tab_Totais.Tab = 0
vTipoEdicaoNFe = "" 'desativei para ver

Me.Left = (Tela_Principal.ScaleWidth - Me.Width) / 2
Me.Top = (Tela_Principal.ScaleHeight - Me.Height) / 2

cboIndicadorPagamento.Clear
cboIndicadorPagamento.AddItem "0 - Pagamento à vista", 0
cboIndicadorPagamento.AddItem "1 - Pagamento à prazo", 1
cboIndicadorPagamento.AddItem "2 - Outros", 2

cboFormatoDANFe.Clear
cboFormatoDANFe.AddItem "1 - Retrato", 0
cboFormatoDANFe.AddItem "2 - Paisagem", 1

cboTipoEmissao.Clear
cboTipoEmissao.AddItem "1 - Normal", 0
cboTipoEmissao.AddItem "2 - Contingência FS", 1
cboTipoEmissao.AddItem "3 - Contingência SCAN", 2
cboTipoEmissao.AddItem "4 - Contingência DPEC", 3
cboTipoEmissao.AddItem "5 - Contingência FS-DA", 4
cboTipoEmissao.AddItem "6 - Contingência SVC-AN", 5

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmNota.Enabled = False
frmDestinatario.Enabled = False
Tab_Totais.Enabled = False
Tab_Produtos.Enabled = False 'desabilitei por causa do pedro caruaru
cmdDuplicar.Enabled = False
'frmTransmissao.Enabled = False
frmItens.Enabled = False
TipoSelecaoConsulta = "0"
TipoSelecaoConsulta = 0

cboFiltroNota_GotFocus
cboFiltroNota.ListIndex = 4
cboConNotaMes.Text = Format(Date, "mmmm")
cboConNotaAno.Text = Year(Date)
cmdExibirConNotas_Click
'ExibirUltimasNfe

sSQL = "SELECT CRT, ESTADO, RegimeTributario, IPICompoeDIFAL FROM empresa"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    vTipoCRT = r("CRT")
    vUFEmpresa = r("ESTADO")
    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))
    vIPICompoeDIFAL = IIf(IsNull(r("IPICompoeDIFAL")), 0, r("IPICompoeDIFAL"))
    If Left(cboDestOperacao.Text, 1) = 2 Then
        vAliqUFInter = Format(12, "#0.00")
        vAliqUFDest = Format(18, "#0.00")
    Else
        vAliqUFInter = Format(0, "#0.00")
        vAliqUFDest = Format(0, "#0.00")
    End If
End If
End Sub

Private Sub GridNotas_DblClick()
'Clear_Controls
'LimparObjetosProduto
'If cmdSalvar.Enabled = True Then
'    MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub
'Else
'    RsOpen TbNotas, "SELECT *,  " & _
'                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE (CASE WHEN Inutilizada = 1 THEN 'Inutilizada' ELSE 'Em Digitação' END) END) END) END) AS Status " & _
'                    "FROM NotaFiscal WHERE CodigoNota = " & GridNotas.TextMatrix(GridNotas.Row, 1)
'    Load_Controls
'    Frm_NF.Tab = 0
'End If
End Sub

Private Sub GridNotasItens_Click()
Dim bEditavel As Boolean
bEditavel = False

Select Case GridNotasItens.Col
    Case 2, 5, 6, 7, 8, 17, 19, 21, 23, 24, 25
        bEditavel = True
End Select

If bEditavel And GridNotasItens.Row > 0 And GridNotasItens.TextMatrix(GridNotasItens.Row, 1) <> "" Then
    txtEdit.Move GridNotasItens.Left + GridNotasItens.CellLeft, GridNotasItens.Top + GridNotasItens.CellTop, GridNotasItens.CellWidth, GridNotasItens.CellHeight
    txtEdit.Text = GridNotasItens.TextMatrix(GridNotasItens.Row, GridNotasItens.Col)
    txtEdit.Visible = True
    txtEdit.SetFocus
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
    iRow = GridNotasItens.Row
    iCol = GridNotasItens.Col
End If
End Sub

Private Sub Label26_Click()
'chkDesc.Value = 1
'chkDesc_Click
End Sub

Private Sub lblCodFabrica_Click()
chkCodBarra.Value = 1
'chkCodBarra_Click
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

Private Sub mskFinalPedidos_GotFocus()
SelectControl mskFinalPedidos
End Sub

Private Sub mskFinalPedidos_KeyPress(KeyAscii As Integer)
mskFinalPedidos.Mask = "##/##/##"
End Sub


Private Sub mskFinalPedidos_LostFocus()
If mskFinalPedidos.Text = "" Or mskFinalPedidos.Text = "__/__/__" Then
   mskFinalPedidos.Mask = ""
   mskFinalPedidos.Text = ""
   Exit Sub
Else
   If IsDate(mskFinalPedidos.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFinalPedidos.SetFocus
      SelectControl mskFinalPedidos
   End If
End If
End Sub

Private Sub mskInicialPedidos_GotFocus()
SelectControl mskInicialPedidos
End Sub

Private Sub mskInicialPedidos_KeyPress(KeyAscii As Integer)
mskInicialPedidos.Mask = "##/##/##"
End Sub


Private Sub mskInicialPedidos_LostFocus()
If mskInicialPedidos.Text = "" Or mskInicialPedidos.Text = "__/__/__" Then
   mskInicialPedidos.Mask = ""
   mskInicialPedidos.Text = ""
   Exit Sub
Else
   If IsDate(mskInicialPedidos.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskInicialPedidos.SetFocus
      SelectControl mskInicialPedidos
   End If
End If
End Sub



Private Sub txtFrete_Change()
If txtFrete.Text = "" Then Exit Sub
Calcular_TotalItem
End Sub

Private Sub txtFrete_GotFocus()
SelectControl txtFrete
End Sub


Private Sub txtFrete_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtFrete_LostFocus()
If txtFrete.Text = "" Then txtFrete.Text = "0"
txtFrete.Text = Format(txtFrete.Text, ocMONEY)
Calcular_TotalItem
End Sub


Private Sub txtInfAdicionais_GotFocus()
txtInfAdicionais.SelStart = 0
txtInfAdicionais.SelLength = Len(txtInfAdicionais)
End Sub


Private Sub txtIntervaloDup_GotFocus()
SelectControl txtIntervaloDup
End Sub


Private Sub txtIntervaloDup_LostFocus()
Calcular_Prazo
End Sub


Private Sub txtNatureza_GotFocus()
SelectControl txtNatureza
End Sub


Private Sub txtNumDup_GotFocus()
SelectControl txtNumDup
End Sub


Private Sub txtNumNota_Change()
'MsgBox txtNumNota.Text
End Sub

Private Sub txtNumParcDup_Change()
Calcular_Parcelas
End Sub

Private Sub txtNumParcDup_GotFocus()
SelectControl txtNumParcDup
End Sub


Private Sub txtNumParcDup_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtNumParcDup_LostFocus()
Calcular_Parcelas
End Sub


Private Sub txtOutrosItem_Change()
If txtOutrosItem.Text = "" Then Exit Sub
Calcular_TotalItem
End Sub

Private Sub txtOutrosItem_GotFocus()
SelectControl txtOutrosItem
End Sub


Private Sub txtOutrosItem_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtOutrosItem_LostFocus()
If txtOutrosItem.Text = "" Then txtOutrosItem.Text = "0"
txtOutrosItem.Text = Format(txtOutrosItem.Text, ocMONEY)
Calcular_TotalItem
End Sub


Private Sub txtQuant_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtSeguro_Change()
If txtSeguro.Text = "" Then Exit Sub
Calcular_TotalItem
End Sub

Private Sub txtSeguro_GotFocus()
SelectControl txtSeguro
End Sub


Private Sub txtSeguro_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


Private Sub txtSeguro_LostFocus()
If txtSeguro.Text = "" Then txtSeguro.Text = "0"
txtSeguro.Text = Format(txtSeguro.Text, ocMONEY)
Calcular_TotalItem
End Sub


Private Sub txtTotalDup_GotFocus()
SelectControl txtTotalDup
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtValorDesconto_GotFocus()
SelectControl txtValorDesconto
End Sub

Private Sub txtValorDesconto_LostFocus()
If txtValorDesconto.Text = "" Then txtValorDesconto.Text = "0"
Moeda txtValorDesconto
Call DistribuirDesconto
AtualizarTotaisNota
End Sub


Private Sub txtValorICMS_GotFocus()
SelectControl txtValorICMS
End Sub

Private Sub txtValorICMS_LostFocus()
Moeda txtValorICMS
AtualizarTotaisNota
End Sub


Private Sub txtValorICMSST_GotFocus()
SelectControl txtValorICMSST
End Sub

Private Sub txtValorICMSST_LostFocus()
If txtValorICMSST.Text = "" Then txtValorICMSST.Text = "0"
Moeda txtValorICMSST
AtualizarTotaisNota
End Sub


Private Sub txtValorIPI_GotFocus()
SelectControl txtValorIPI
End Sub

Private Sub txtValorIPI_LostFocus()
If txtValorIPI.Text = "" Then txtValorIPI.Text = "0"
Moeda txtValorIPI
AtualizarTotaisNota
End Sub


Private Sub txtVolQuant_GotFocus()
txtVolQuant.SelStart = 0
txtVolQuant.SelLength = Len(txtVolQuant)
End Sub

Private Sub txtVolQuant_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtVolEspecie_GotFocus()
txtVolEspecie.SelStart = 0
txtVolEspecie.SelLength = Len(txtVolEspecie)
End Sub


Private Sub txtVolMarca_GotFocus()
txtVolMarca.SelStart = 0
txtVolMarca.SelLength = Len(txtVolMarca)
End Sub


Private Sub txtVolNumeracao_GotFocus()
txtVolNumeracao.SelStart = 0
txtVolNumeracao.SelLength = Len(txtVolNumeracao)
End Sub


Private Sub txtCodObservacao_GotFocus()
txtCodObservacao.SelStart = 0
txtCodObservacao.SelLength = Len(txtCodObservacao)
End Sub

Private Sub txtValorOutrasDespesas_GotFocus()
SelectControl txtValorOutrasDespesas
End Sub

Private Sub txtVolPesoBruto_GotFocus()
txtVolPesoBruto.SelStart = 0
txtVolPesoBruto.SelLength = Len(txtVolPesoBruto)
End Sub

Private Sub txtVolPesoLiquido_GotFocus()
txtVolPesoLiquido.SelStart = 0
txtVolPesoLiquido.SelLength = Len(txtVolPesoLiquido)
End Sub

Private Sub txtValorSeguro_GotFocus()
SelectControl txtValorSeguro
End Sub

Private Sub Mostrar_Pedido(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then

Dim totalRegistros As Long

    'buscar Numero e codigo da nota (autopreenchimento)
    RsOpen TbNotaPedido, "SELECT CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota, SerieNF,  " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM NotaFiscal"
                
    If TbNotaPedido.RecordCount > 0 Then totalRegistros = TbNotaPedido.RecordCount
        
    'Clear_Controls
    
    'INICIO DO PREENCHIMENTOS DOS OBJETOS
    If TbNotaPedido.EOF And TbNotaPedido.BOF Then
        txtNumNota.Text = "1"
        txtCodNota.Text = "1"
        txtSerie.Text = "1"
    Else
        TbNotaPedido.MoveLast
        txtNumNota.Text = TbNotaPedido("NumeroNota") + 1
        txtCodNota.Text = TbNotaPedido("CodigoNota") + 1
        txtSerie.Text = TbNotaPedido("SerieNF")
    End If

    txtCodCliente = Format(rTabela("COD_CLIENTE"), "@")
    cboCliente = Format(rTabela("varnome"), "@")

    cboDestOperacao.Text = "1 - Operação Interna"
    'txtInfAdicionais = Format(rTabela("InformacoesAdicionais"), "@")
    cboNatureza = "5102"
    txtNatureza = "VENDA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS"
    cboFinalidade = "1 - NFe NORMAL"
    cboTipoDest = "CLIENTE"
    'mskEmissao = Format(Date, "dd/mm/yyyy")
    'mskSaida = Format(Date, "dd/mm/yyyy")
    'mskHora = Format(Time(), "HH:MM:ss")
    txtVolPesoBruto = Format(0, "@")
    txtVolPesoLiquido = Format(0, "@")
    'txtPlacaUF = Format(0, "@")
    cboModFrete = "9 - SEM FRETE"
    'txtCodTransporte = Format(0, "@")
    'cboTransporte = Format(0, "@")
    'txtPlaca = Format(0, "@")
    txtVolQuant = Format(0, "@")
    txtVolEspecie = Format(0, "@")
    txtVolMarca = Format(0, "@")
    txtVolNumeracao = Format(0, "@")
    'txtCodObservacao = Format(0, "@")
    
    txtValorSeguro = Format(0, "##,##0.00")
    txtValorOutrasDespesas = Format(0, "##,##0.00")
    txtValorFrete = Format(0, "##,##0.00")
    txtBaseICMS = Format(0, "##,##0.00")
    txtBaseICMSST = Format(0, "##,##0.00")
    txtValorIPI = Format(0, "##,##0.00")
    txtValorICMS = Format(0, "##,##0.00")
    txtValorICMSST = Format(0, "##,##0.00")
    txtValorDesconto = Format(0, "##,##0.00")
    txtTotaldaNota = FormatNumber(0, 2)
    txtTotaldosProdutos = FormatNumber(0, 2)
    
    'transmissão
    'Text30 = Format(0, "@")
    'Text31 = Format(0, "@")
    'Text32 = Format(0, "@")
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
End If
End Sub

Private Sub txtDesc_Change()
If txtDesc.Text = "" Or txtValor.Text = "" Then Exit Sub
Calcular_TotalItem
End Sub

Private Sub Calcular_TotalItem()
If txtCodProduto = "" Then Exit Sub
If txtValor.Text = "" Or IsNumeric(txtValor.Text) = False Then Exit Sub

Dim varValor    As Currency
Dim varQuant    As Double
Dim varFrete    As Currency
Dim varSeguro   As Currency
Dim varOutros   As Currency
Dim varDesc     As Currency
Dim varSubtotal As Currency

varValor = CCur(txtValor.Text)
varQuant = CDbl(IIf(txtQuant.Text = "", "1", txtQuant.Text))
varFrete = CCur(IIf(txtFrete.Text = "", "0", txtFrete.Text))
varSeguro = CCur(IIf(txtSeguro.Text = "", "0", txtSeguro.Text))
varOutros = CCur(IIf(txtOutrosItem.Text = "", "0", txtOutrosItem.Text))
varDesc = CCur(IIf(txtDesc.Text = "", "0", txtDesc.Text))

' Base = (Valor x Quant + Frete + Seguro + Outros) - Desconto
varSubtotal = (varValor * varQuant) + varFrete + varSeguro + varOutros - varDesc

txtSubTotal.Text = FormatNumber(varSubtotal, 2)
End Sub

Private Sub txtDesc_GotFocus()
SelectControl txtDesc
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtDesc_LostFocus()
On Error GoTo erro
If txtDesc.Text = "" Then txtDesc.Text = "0"
txtDesc.Text = Format(txtDesc.Text, ocMONEY)
Calcular_TotalItem
Exit Sub

erro:
   ShowMsg "O valor digitado é inválido!", vbExclamation
   txtDesc.Text = 0
End Sub

Private Sub txtnumnota_GotFocus()
txtNumNota.SelStart = 0
txtNumNota.SelLength = Len(txtNumNota)
End Sub

Private Sub txtBaseICMSST_GotFocus()
SelectControl txtBaseICMSST
End Sub

Private Sub txtBaseICMS_GotFocus()
txtBaseICMS.SelStart = 0
txtBaseICMS.SelLength = Len(txtBaseICMS)
End Sub

Private Sub txtPlacaUF_GotFocus()
txtPlacaUF.SelStart = 0
txtPlacaUF.SelLength = Len(txtPlacaUF)
End Sub

Private Sub txtPlacaUF_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtPlaca_GotFocus()
txtPlaca.SelStart = 0
txtPlaca.SelLength = Len(txtPlaca)
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtQuant_Validate(Cancel As Boolean)
Calcular_TotalItem
End Sub

Private Sub txtSubTotal_GotFocus()
SelectControl txtSubTotal
End Sub


Private Sub txtValor_GotFocus()
SelectControl txtValor
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
Calcular_TotalItem
End Sub

Private Sub txtValorFrete_GotFocus()
SelectControl txtValorFrete
End Sub

Private Sub txtValorFrete_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtValorFrete_LostFocus()
If txtValorFrete.Text = "" Then txtValorFrete.Text = "0"
Moeda txtValorFrete
Call DistribuirFrete
AtualizarTotaisNota
End Sub

Private Sub txtCodObservacao_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbMSN As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If txtCodObservacao.Text = "" Then
        'Frm_LMensagem.Show 1
        Exit Sub
    End If
    RsOpen TbMSN, "select * from ObservacoesNFe where CodigoObservacao = " & Val(txtCodObservacao.Text)
    If TbMSN.EOF And TbMSN.BOF Then
        MsgBox "Não foi possivel localizar o código da observação solicitada.", vbCritical, "Online Commerce": Exit Sub
    Else
        txtCodObservacao.Text = TbMSN("CodigoObservacao")
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtCodObservacao_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtValorOutrasDespesas_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub


End Sub

Private Sub txtValorOutrasDespesas_LostFocus()
If txtValorOutrasDespesas.Text = "" Then txtValorOutrasDespesas.Text = "0"
Moeda txtValorOutrasDespesas
Call DistribuirOutros
AtualizarTotaisNota
End Sub

Private Sub txtVolPesoBruto_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub


End Sub

Private Sub txtVolPesoBruto_LostFocus()
If txtVolPesoBruto.Text = "" Then txtVolPesoBruto.Text = "0"
End Sub

Private Sub txtVolPesoLiquido_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtVolPesoLiquido_LostFocus()
If txtVolPesoLiquido.Text = "" Then txtVolPesoLiquido.Text = "0"
End Sub

Private Sub txtValorSeguro_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtValorSeguro_LostFocus()
If txtValorSeguro.Text = "" Then txtValorSeguro.Text = "0"
Moeda txtValorSeguro
Call DistribuirSeguro
AtualizarTotaisNota
End Sub

Private Sub cboDestOperacao_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim TbOP As New ADODB.Recordset
'On Error GoTo erro
'If KeyCode = 13 Then
'    If cboDestOperacao.Text = "" Then
'        cboDestOperacao.SetFocus
'        Exit Sub
'    End If
'    RsOpen TbOP, "SELECT * FROM NaturezaOperacaoNF WHERE CodigoNatureza = " & cboDestOperacao.Text
'    If TbOP.EOF And TbOP.BOF Then
        'MsgBox "Não foi possivel localizar o código da OP no sistema.", vbCritical, "Online Commerce": Exit Sub
'    Else
'        cboNatureza.Text = TbOP("descricao")
'        cboNatureza.SetFocus
'    End If
'End If
'Exit Sub
'erro:
'MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub
Private Sub txtnumnota_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtBaseICMSST_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtBaseICMSST_LostFocus()
If txtBaseICMSST.Text = "" Then txtBaseICMSST.Text = "0"
Moeda txtBaseICMSST
AtualizarTotaisNota
End Sub

Private Sub txtBaseICMS_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtBaseICMS_LostFocus()
If txtBaseICMS.Text = "" Then txtBaseICMS.Text = "0"
Moeda txtBaseICMS
AtualizarTotaisNota
End Sub

Private Sub txtCodBarra_GotFocus()
SelectControl txtCodBarra
End Sub


Private Sub txtCodBarra_LostFocus()
If TipoSelecaoConsulta = "0" Or TipoSelecaoConsulta = "1" Then
    If txtCodBarra.Text = "" Then
        TipoSelecaoConsulta = "0"
        LimparObjetosProduto
    Else
        sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante, quant_estoque, unid_medida, CFOP, NCM, ICMSCST, ICMSAliq, EAN  FROM produtos WHERE (COD_BARRA = '" & txtCodBarra.Text & "') AND (ativo = 1);"
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r.BOF Then
           txtCodProduto.Text = r("var_codprod")
           
          ' If tipoEmpresa = 4 Then
           '    cboDescricao.Text = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante")) & " /  " & r("REF")
            '   'cboDescricao2.Text = ValidateNull(r("var_desc"))
           '     MostrarValorVenda
            '    txtQuant.SetFocus
           'Else
              'txtEAN.Text = ValidateNull(r("EAN"))
              'cboDescricao.Text = ValidateNull(r("var_desc"))
              'txtUnid.Text = ValidateNull(r("unid_medida"))
              'txtCFOP.Text = ValidateNull(r("CFOP"))
              'txtCST.Text = ValidateNull(r("ICMSCST"))
              'txtNCM.Text = ValidateNull(r("NCM"))
              'txtICMS.Text = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
            TipoSelecaoConsulta = "1"
            MostrarValorVenda
            Mostrar_Aliquotas_Produto
            txtQuant.SetFocus
            
           'End If

            
            'cboDescricao.Locked = True
            
           ' MostrarValorVenda
            'txtQuant.SetFocus
        Else
           ShowMsg "Produto Inexistente!", vbCritical
           TipoSelecaoConsulta = "0"
           LimparObjetosProduto
           txtCodBarra.SetFocus
           'Exit Sub
        End If
    End If
End If

If TipoSelecaoConsulta = "1" Then
    txtCodBarra.BackColor = &HC0FFFF
    cboDescricao.BackColor = &HFFFFFF
    cboDescricao.Locked = True
ElseIf TipoSelecaoConsulta = "2" Then
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HC0FFFF
    txtCodBarra.Locked = True
Else
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HFFFFFF
    txtCodBarra.Locked = False
    cboDescricao.Locked = False
End If
On Local Error Resume Next
End Sub

Private Sub txtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbClientes As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If txtCodCliente.Text = "" Then
        txtCodCliente.SetFocus
        Exit Sub
    End If
    RsOpen TbClientes, "SELECT * FROM cliente WHERE codigo = " & Val(txtCodCliente.Text)
    If TbClientes.EOF And TbClientes.BOF Then
        MsgBox "Código do cliente não foi localizado no sistema. Verifique.", vbCritical, "Online Commerce": Exit Sub
    Else
        cboCliente.Text = TbClientes("nome")
        txtCliEndereco.Text = TbClientes("nome")
        txtCliNum.Text = TbClientes("nome")
        txtCliBairro.Text = TbClientes("nome")
        txtCliCidade.Text = TbClientes("nome")
        txtCliUF.Text = TbClientes("nome")
        txtCliCPF.Text = TbClientes("nome")
        txtCliIE.Text = TbClientes("nome")
        
        cboCliente.SetFocus
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNatureza_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub mskEmissao_GotFocus()
If mskEmissao.Text = "" Then mskEmissao = Format(Date, "dd/mm/yyyy")

mskEmissao.SelStart = 0
mskEmissao.SelLength = Len(mskEmissao)
End Sub

Private Sub mskSaida_GotFocus()
If mskSaida.Text = "" Then mskSaida = Format(Date, "dd/mm/yyyy")
mskSaida.SelStart = 0
mskSaida.SelLength = Len(mskSaida)
End Sub

Private Sub mskHora_GotFocus()
If mskHora.Text = "" Then mskHora = Format(Time(), "HH:MM:ss")
mskHora.SelStart = 0
mskHora.SelLength = Len(mskHora)
End Sub

Private Sub txtCodNota_Change()
Mostrar_ItensNota
MostrarCorrecao
End Sub

Private Sub txtCodTransporte_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Function MostraStatus_F9() As String
 On Error Resume Next
 xCancelada = TbNotas("Cancelada")
 If TbNotas("Denegada") Or TbNotas("Inutilizada") Then
    MostraStatus.ForeColor = vbRed
    MostraStatus_F9$ = "Denegada"  'Deve retornar uma expressão caractere
    If TbNotas("Inutilizada") Then MostraStatus_F9$ = "Inutilizada"
 ElseIf TbNotas("Enviada") And Not TbNotas("Cancelada") Then
    MostraStatus.ForeColor = vbBlue
    MostraStatus_F9$ = "Transmitida/Autorizada"  'Deve retornar uma expressão caractere
 ElseIf TbNotas("Cancelada") Then
    MostraStatus.ForeColor = vbRed
    MostraStatus_F9$ = "Cancelada"  'Deve retornar uma expressão caractere
 Else
    MostraStatus.ForeColor = vbBlack
    MostraStatus_F9$ = "Em Digitação"  'Deve retornar uma expressão caractere
 End If
End Function


Private Sub LimparGridNotas()
   Dim i As Integer
   
   With GridNotas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 500
      .ColWidth(2) = 1000
      .ColWidth(3) = 1100
      .ColWidth(4) = 2000
      .ColWidth(5) = 4000
      .ColWidth(6) = 1500
      .ColWidth(7) = 1100
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "NUM NF"
      .TextMatrix(0, 3) = "DATA"
      .TextMatrix(0, 4) = "NATUREZA OP"
      .TextMatrix(0, 5) = "DESTINATÁRIO"
      .TextMatrix(0, 6) = "VALOR NOTA"
      .TextMatrix(0, 7) = "STATUS"
      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 2
      .ColAlignment(7) = 1
      .rows = .rows + 1
      
      i = i + 1
      .rows = .rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'GridNotas.ColWidth(0) = 400
      'GridNotas.Rows = 11
      GridNotas.Col = 0
      
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub LimparGridItensNota()
Dim i As Integer, j As Integer

With GridNotasItens
   .Visible = False
   .Redraw = False
   
   .Clear
   .Cols = 17
   .rows = 2
   
   .ColWidth(0) = 200
   .ColWidth(1) = 400
   .ColWidth(2) = 1500
   .ColWidth(3) = 0
   .ColWidth(4) = 3500
   .ColWidth(5) = 450 '
   .ColWidth(6) = 900 '500
   .ColWidth(7) = 600
   .ColWidth(8) = 500
   .ColWidth(9) = 850
   .ColWidth(10) = 850
   .ColWidth(11) = 800
   .ColWidth(12) = 850
   .ColWidth(13) = 700
   .ColWidth(14) = 850
   .ColWidth(15) = 0
   .ColWidth(16) = 0

   .TextMatrix(0, 1) = "No."
   .TextMatrix(0, 2) = "EAN"
   .TextMatrix(0, 3) = "CÓD."
   .TextMatrix(0, 4) = "DESCRIÇÃO"
   .TextMatrix(0, 5) = "UND"
   .TextMatrix(0, 6) = "NCM"
   .TextMatrix(0, 7) = "CFOP"
   .TextMatrix(0, 8) = "CST"
   .TextMatrix(0, 9) = "ALIQ."
   .TextMatrix(0, 10) = "ICMS"
   .TextMatrix(0, 11) = "VALOR"
   .TextMatrix(0, 12) = "QTDE"
   .TextMatrix(0, 13) = "DESC."
   .TextMatrix(0, 14) = "TOTAL"
   .TextMatrix(0, 15) = "IPI"
   .TextMatrix(0, 16) = "IPI"
   
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
   
   i = 1
   
   'ALINHAMENTO
   .ColAlignment(0) = 1
   .ColAlignment(1) = 1
   .ColAlignment(2) = 1
   .ColAlignment(3) = 1
   .ColAlignment(4) = 1
   .ColAlignment(5) = 1
   .ColAlignment(6) = 1
   .ColAlignment(7) = 1
   .ColAlignment(8) = 1
   .ColAlignment(9) = 6
   .ColAlignment(10) = 6
   .ColAlignment(11) = 6
   .ColAlignment(12) = 6
   .ColAlignment(13) = 6
   .ColAlignment(14) = 6
   .ColAlignment(15) = 6
   .ColAlignment(16) = 6
   
   GridNotasItens.Col = 0
         
   .Visible = True
   .Redraw = True
End With
End Sub


Public Sub FormatarGridNotas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridNotas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 15
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 900
      .ColWidth(3) = 1100
      .ColWidth(4) = 1350
      .ColWidth(5) = 4000
      .ColWidth(6) = 1100
      .ColWidth(7) = 1100
      .ColWidth(8) = 3300
      .ColWidth(9) = 1600
      .ColWidth(10) = 1600
      .ColWidth(11) = 0
      .ColWidth(12) = 0
      .ColWidth(13) = 0
      .ColWidth(14) = 900
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "NUM NF"
      .TextMatrix(0, 3) = "DATA"
      .TextMatrix(0, 4) = "FINALIDADE"
      .TextMatrix(0, 5) = "DESTINATÁRIO"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "STATUS"
      .TextMatrix(0, 8) = "CHAVE"
      .TextMatrix(0, 9) = "RECIBO"
      .TextMatrix(0, 10) = "PROTOCOLO"
      .TextMatrix(0, 11) = "DATA PROT."
      .TextMatrix(0, 12) = "CÓD."
      .TextMatrix(0, 13) = "TIPO"
      .TextMatrix(0, 14) = "SÉRIE"
      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("CodigoNota")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("NumeroNota"), "000000")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("DataEmissao"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 4) = rTabela("FinalidadeEmissaoNFe")
            .TextMatrix(.rows - 1, 5) = rTabela("RazaoSocial")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("ValorNota"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = rTabela("Status")
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("ChavedeAcesso"))
            .TextMatrix(.rows - 1, 9) = rTabela("NumeroRecibo")
            .TextMatrix(.rows - 1, 10) = rTabela("NumeroProtocolo")
            .TextMatrix(.rows - 1, 11) = rTabela("DataHoraProcotolo")
            .TextMatrix(.rows - 1, 12) = rTabela("CodigoCorrentista")
            .TextMatrix(.rows - 1, 13) = ValidateNull(rTabela("TipoCliente"))
            .TextMatrix(.rows - 1, 14) = rTabela("SERIENF")
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      .rows = .rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
              
     'GridNotas.ColWidth(0) = 400
      'GridNotas.Rows = 11
      GridNotas.Col = 0
            
      .Visible = True
      .Redraw = True
   End With
End Sub
Private Sub txtQuant_Change()
If txtQuant.Text = "" Or txtValor.Text = "" Then Exit Sub
Calcular_TotalItem
End Sub

Private Sub txtQuant_GotFocus()
SelectControl txtQuant
End Sub


Private Sub txtQuant_LostFocus()
If txtQuant.Text = "" Then txtQuant.Text = "1"
txtQuant.Text = Format(txtQuant.Text, ocPESO)
Calcular_TotalItem
End Sub


Private Sub txtValor_Change()
If txtValor.Text = "" Then Exit Sub
Calcular_TotalItem
End Sub


Private Sub txtValor_LostFocus()
If txtValor.Text = "" Then Exit Sub
txtValor.Text = Format(txtValor.Text, ocMONEY)
Calcular_TotalItem
End Sub


