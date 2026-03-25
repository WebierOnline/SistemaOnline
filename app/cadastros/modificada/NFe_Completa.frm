VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form NFe_Completa 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NFe - Nota Fiscal Eletronica"
   ClientHeight    =   10590
   ClientLeft      =   735
   ClientTop       =   1455
   ClientWidth     =   15510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Frm_NF"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10590
   ScaleWidth      =   15510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   0
      TabIndex        =   67
      Top             =   840
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   16536
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   5292
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "NFe_Completa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdConsultar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdRecalcularNF"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelarNota"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdTransmitir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdImprimir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSair"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAlterar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSalvar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdExcluir"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdNovo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCancelar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SSTab3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "SSTab2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "frmDestinatario"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCodObservacao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "frmNota"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "NOTAS FISCAIS"
      TabPicture(1)   =   "NFe_Completa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "GridNotas"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "PEDIDOS"
      TabPicture(2)   =   "NFe_Completa.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame4 
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   181
         Top             =   8100
         Width           =   13155
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
            Height          =   735
            Left            =   4320
            TabIndex        =   185
            Top             =   240
            Width           =   5535
            Begin VB.TextBox txtConNotaNumNota 
               Height          =   315
               Left            =   1200
               TabIndex        =   190
               Top             =   300
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtConNotaCodCliente 
               Height          =   315
               Left            =   4740
               TabIndex        =   189
               Top             =   300
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox cboConNotaCliente 
               Height          =   315
               Left            =   840
               TabIndex        =   188
               Top             =   300
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.ComboBox cboConNotaAno 
               Height          =   315
               Left            =   600
               TabIndex        =   187
               Top             =   300
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboConNotaMes 
               Height          =   315
               Left            =   2580
               TabIndex        =   186
               Top             =   300
               Visible         =   0   'False
               Width           =   1155
            End
            Begin MSMask.MaskEdBox mskConNotaFinal 
               Height          =   315
               Left            =   2760
               TabIndex        =   191
               Top             =   300
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal2 
               Height          =   315
               Left            =   3780
               TabIndex        =   192
               Top             =   300
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
            Begin MSMask.MaskEdBox mskConNotaInicial 
               Height          =   315
               Left            =   720
               TabIndex        =   193
               Top             =   300
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal1 
               Height          =   315
               Left            =   1740
               TabIndex        =   194
               Top             =   300
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
               MICON           =   "NFe_Completa.frx":2452
               PICN            =   "NFe_Completa.frx":246E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblConNotaNumNota 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Num. Nota:"
               Height          =   195
               Left            =   180
               TabIndex        =   200
               Top             =   360
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label lblConNotaFinal 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Final:"
               Height          =   195
               Left            =   2280
               TabIndex        =   199
               Top             =   360
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblConNotaInicial 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   198
               Top             =   360
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lblConNotaCliente 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   180
               TabIndex        =   197
               Top             =   360
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblConNotaMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   2160
               TabIndex        =   196
               Top             =   360
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblConNotaAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   180
               TabIndex        =   195
               Top             =   360
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
            Height          =   735
            Left            =   180
            TabIndex        =   182
            Top             =   240
            Width           =   4035
            Begin VB.ComboBox cboFiltroNota 
               Height          =   315
               Left            =   960
               TabIndex        =   183
               Top             =   300
               Width           =   2715
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Escolher:"
               Height          =   195
               Left            =   180
               TabIndex        =   184
               Top             =   360
               Width           =   660
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConNotas 
            Height          =   495
            Left            =   9960
            TabIndex        =   201
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
            MICON           =   "NFe_Completa.frx":4850
            PICN            =   "NFe_Completa.frx":486C
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
         Height          =   1635
         Left            =   120
         TabIndex        =   166
         Top             =   360
         Width           =   13275
         Begin VB.TextBox Text30 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2280
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1200
            Width           =   4215
         End
         Begin VB.TextBox Text31 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6540
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1635
         End
         Begin VB.TextBox Text32 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   8220
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1635
         End
         Begin VB.ComboBox cboFinalidade 
            Height          =   315
            ItemData        =   "NFe_Completa.frx":5146
            Left            =   2400
            List            =   "NFe_Completa.frx":5148
            TabIndex        =   3
            Top             =   540
            Width           =   1995
         End
         Begin VB.ComboBox cboTipoNota 
            Height          =   315
            Left            =   900
            TabIndex        =   2
            Top             =   540
            Width           =   1455
         End
         Begin VB.TextBox txtNatureza 
            Height          =   315
            Left            =   7620
            TabIndex        =   6
            Top             =   540
            Width           =   4260
         End
         Begin VB.ComboBox cboNatureza 
            Height          =   315
            Left            =   6600
            TabIndex        =   5
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtNumNota 
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   1
            Top             =   540
            Width           =   720
         End
         Begin VB.ComboBox cboDestOperacao 
            Height          =   315
            Left            =   4440
            TabIndex        =   4
            Top             =   540
            Width           =   2115
         End
         Begin ChamaleonBtn.chameleonButton cmdCal2 
            Height          =   315
            Left            =   1080
            TabIndex        =   167
            Tag             =   "Calendario"
            Top             =   1200
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
            MICON           =   "NFe_Completa.frx":514A
            PICN            =   "NFe_Completa.frx":5166
            PICH            =   "NFe_Completa.frx":74B9
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
            Left            =   12900
            TabIndex        =   168
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
            MICON           =   "NFe_Completa.frx":980C
            PICN            =   "NFe_Completa.frx":9828
            PICH            =   "NFe_Completa.frx":BB7B
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
            Left            =   1440
            TabIndex        =   9
            Top             =   1200
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSaida 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEmissao 
            Height          =   315
            Left            =   11940
            TabIndex        =   7
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave de Acesso"
            Height          =   195
            Left            =   2280
            TabIndex        =   179
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Protocolo"
            Height          =   195
            Left            =   6540
            TabIndex        =   178
            Top             =   960
            Width           =   675
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo"
            Height          =   195
            Left            =   8220
            TabIndex        =   177
            Top             =   960
            Width           =   510
         End
         Begin VB.Label MostraStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   9900
            TabIndex        =   13
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finalidade da Emissão"
            Height          =   195
            Left            =   2400
            TabIndex        =   176
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Nota"
            Height          =   195
            Left            =   900
            TabIndex        =   175
            Top             =   300
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NF Num."
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   174
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destino"
            Height          =   195
            Index           =   0
            Left            =   4485
            TabIndex        =   173
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Natureza da Operação"
            Height          =   195
            Left            =   6600
            TabIndex        =   172
            Top             =   300
            Width           =   1620
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Emissão"
            Height          =   195
            Left            =   11940
            TabIndex        =   171
            Top             =   285
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Saida"
            Height          =   195
            Left            =   120
            TabIndex        =   170
            Top             =   945
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
            Height          =   195
            Left            =   1440
            TabIndex        =   169
            Top             =   945
            Width           =   345
         End
      End
      Begin VB.TextBox txtCodObservacao 
         Height          =   315
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   74
         Top             =   8280
         Width           =   1635
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
         Height          =   1335
         Left            =   120
         TabIndex        =   68
         Top             =   2100
         Width           =   13275
         Begin VB.ComboBox cboTipoDest 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox cboConsumidorFinal 
            Height          =   315
            Left            =   10020
            TabIndex        =   17
            Top             =   480
            Width           =   1875
         End
         Begin VB.ComboBox cboTipoContribuinte 
            Height          =   315
            Left            =   7440
            TabIndex        =   16
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox TxtCodCliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            MaxLength       =   10
            TabIndex        =   69
            Top             =   240
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.ComboBox CboCliente 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2220
            TabIndex        =   15
            Top             =   480
            Width           =   5175
         End
         Begin VB.TextBox txtCliIBGE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   6390
            TabIndex        =   23
            ToolTipText     =   "Cód. IBGE"
            Top             =   900
            Width           =   825
         End
         Begin VB.TextBox txtCliIE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   8700
            TabIndex        =   25
            ToolTipText     =   "IE"
            Top             =   900
            Width           =   1125
         End
         Begin VB.TextBox txtCliCPF 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   7260
            TabIndex        =   24
            ToolTipText     =   "CNPJ/CPF"
            Top             =   900
            Width           =   1425
         End
         Begin VB.TextBox txtCliUF 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   6060
            TabIndex        =   22
            ToolTipText     =   "Estado"
            Top             =   900
            Width           =   315
         End
         Begin VB.TextBox txtCliCidade 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   4800
            TabIndex        =   21
            ToolTipText     =   "Cidade"
            Top             =   900
            Width           =   1245
         End
         Begin VB.TextBox txtCliBairro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3660
            TabIndex        =   20
            ToolTipText     =   "Bairro"
            Top             =   900
            Width           =   1125
         End
         Begin VB.TextBox txtCliNum 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3240
            TabIndex        =   19
            ToolTipText     =   "Número"
            Top             =   900
            Width           =   375
         End
         Begin VB.TextBox txtCliEndereco 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Endereço"
            Top             =   900
            Width           =   3105
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Destinatário"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consumidor Final"
            Height          =   195
            Left            =   10020
            TabIndex        =   72
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Contribuinte"
            Height          =   195
            Left            =   7440
            TabIndex        =   71
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente / Destinatário"
            Height          =   195
            Left            =   2220
            TabIndex        =   70
            Top             =   240
            Width           =   1485
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1635
         Left            =   120
         TabIndex        =   75
         Top             =   7620
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   2884
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   3528
         TabCaption(0)   =   "Totais da Nota"
         TabPicture(0)   =   "NFe_Completa.frx":DECE
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
         Tab(0).Control(18)=   "Text3"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtValorDesconto"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtTotaldaNota"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtTotaldosProdutos"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Outros Tributos"
         TabPicture(1)   =   "NFe_Completa.frx":DEEA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Retenção de Tributos"
         TabPicture(2)   =   "NFe_Completa.frx":DF06
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "Interestadual"
         TabPicture(3)   =   "NFe_Completa.frx":DF22
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin VB.TextBox txtTotaldosProdutos 
            Height          =   315
            Left            =   11580
            MaxLength       =   50
            TabIndex        =   86
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtTotaldaNota 
            Height          =   315
            Left            =   11580
            MaxLength       =   50
            TabIndex        =   85
            Top             =   1200
            Width           =   1560
         End
         Begin VB.TextBox txtValorDesconto 
            Height          =   315
            Left            =   180
            MaxLength       =   50
            TabIndex        =   84
            Top             =   1200
            Width           =   1560
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   6660
            MaxLength       =   50
            TabIndex        =   83
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtValorICMSST 
            Height          =   315
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   82
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtValorICMS 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   81
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtValorOutrasDespesas 
            Height          =   315
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   80
            Top             =   1200
            Width           =   1560
         End
         Begin VB.TextBox txtValorSeguro 
            Height          =   315
            Left            =   3420
            MaxLength       =   50
            TabIndex        =   79
            Top             =   1200
            Width           =   1560
         End
         Begin VB.TextBox txtBaseICMSST 
            Height          =   315
            Left            =   3420
            MaxLength       =   50
            TabIndex        =   78
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtBaseICMS 
            Height          =   315
            Left            =   165
            MaxLength       =   50
            TabIndex        =   77
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtValorFrete 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   76
            Top             =   1200
            Width           =   1560
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total dos Produtos"
            Height          =   195
            Left            =   11580
            TabIndex        =   97
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total da Nota"
            Height          =   195
            Left            =   11580
            TabIndex        =   96
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do Desconto"
            Height          =   195
            Left            =   180
            TabIndex        =   95
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do IPI"
            Height          =   195
            Left            =   6660
            TabIndex        =   94
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS ST"
            Height          =   195
            Left            =   5040
            TabIndex        =   93
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS"
            Height          =   195
            Left            =   1800
            TabIndex        =   92
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do Frete"
            Height          =   195
            Left            =   1800
            TabIndex        =   91
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outras Despesas"
            Height          =   195
            Left            =   5040
            TabIndex        =   90
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do Seguro"
            Height          =   195
            Left            =   3420
            TabIndex        =   89
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base ICMS ST"
            Height          =   195
            Left            =   3420
            TabIndex        =   88
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base ICMS"
            Height          =   195
            Left            =   165
            TabIndex        =   87
            Top             =   360
            Width           =   795
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   4095
         Left            =   120
         TabIndex        =   98
         Top             =   3480
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   7223
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   467
         TabMaxWidth     =   2999
         TabCaption(0)   =   "Produtos"
         TabPicture(0)   =   "NFe_Completa.frx":DF3E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmItens"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Transporte"
         TabPicture(1)   =   "NFe_Completa.frx":DF5A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cboModFrete"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Cobrança"
         TabPicture(2)   =   "NFe_Completa.frx":DF76
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(1)=   "Frame2"
         Tab(2).Control(2)=   "cboIndicadorPagamento"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Informações"
         TabPicture(3)   =   "NFe_Completa.frx":DF92
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSTab4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "DANFe"
         TabPicture(4)   =   "NFe_Completa.frx":DFAE
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cboTipoEmissao"
         Tab(4).Control(1)=   "cboFormatoDANFe"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Exportação e Compra"
         TabPicture(5)   =   "NFe_Completa.frx":DFCA
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         TabCaption(6)   =   "NF Referenciadas"
         TabPicture(6)   =   "NFe_Completa.frx":DFE6
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "txtChaveReferenciada"
         Tab(6).Control(1)=   "Label63"
         Tab(6).ControlCount=   2
         Begin VB.TextBox txtChaveReferenciada 
            Height          =   285
            Left            =   -74820
            MaxLength       =   44
            TabIndex        =   51
            Top             =   600
            Width           =   9555
         End
         Begin VB.Frame Frame3 
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
            Height          =   975
            Left            =   -74880
            TabIndex        =   122
            Top             =   1980
            Width           =   12615
            Begin VB.TextBox txtIntervaloDup 
               Height          =   315
               Left            =   4140
               MaxLength       =   50
               TabIndex        =   44
               Top             =   480
               Width           =   720
            End
            Begin VB.TextBox txtValorParcDup 
               Height          =   315
               Left            =   6195
               MaxLength       =   50
               TabIndex        =   46
               Top             =   480
               Width           =   1320
            End
            Begin VB.TextBox txtTotalDup 
               Height          =   315
               Left            =   1755
               MaxLength       =   50
               TabIndex        =   42
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtNumParcDup 
               Height          =   315
               Left            =   3375
               MaxLength       =   50
               TabIndex        =   43
               Top             =   480
               Width           =   720
            End
            Begin VB.TextBox txtNumDup 
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   41
               Top             =   480
               Width           =   1560
            End
            Begin ChamaleonBtn.chameleonButton chameleonButton1 
               Height          =   315
               Left            =   5880
               TabIndex        =   123
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
               MICON           =   "NFe_Completa.frx":E002
               PICN            =   "NFe_Completa.frx":E01E
               PICH            =   "NFe_Completa.frx":10371
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
               Left            =   4920
               TabIndex        =   45
               Top             =   480
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio:"
               Height          =   195
               Left            =   4920
               TabIndex        =   129
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Intervalo"
               Height          =   195
               Left            =   4140
               TabIndex        =   128
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valoe da Parcela"
               Height          =   195
               Left            =   6195
               TabIndex        =   127
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   195
               Left            =   1755
               TabIndex        =   126
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Parcelas"
               Height          =   195
               Left            =   3375
               TabIndex        =   125
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número/Doc."
               Height          =   195
               Left            =   120
               TabIndex        =   124
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
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
            TabIndex        =   117
            Top             =   1020
            Width           =   12615
            Begin VB.TextBox txtNumFatura 
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   37
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtDescFatura 
               Height          =   315
               Left            =   3375
               MaxLength       =   50
               TabIndex        =   39
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtSubtotalFatura 
               Height          =   315
               Left            =   1755
               MaxLength       =   50
               TabIndex        =   38
               Top             =   480
               Width           =   1560
            End
            Begin VB.TextBox txtTotalFatura 
               Height          =   315
               Left            =   4995
               MaxLength       =   50
               TabIndex        =   40
               Top             =   480
               Width           =   1560
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número"
               Height          =   195
               Left            =   120
               TabIndex        =   121
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto"
               Height          =   195
               Left            =   3375
               TabIndex        =   120
               Top             =   240
               Width           =   690
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal"
               Height          =   195
               Left            =   1755
               TabIndex        =   119
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   195
               Left            =   4995
               TabIndex        =   118
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.ComboBox cboTipoEmissao 
            Height          =   315
            Left            =   -72420
            TabIndex        =   50
            Top             =   720
            Width           =   5055
         End
         Begin VB.ComboBox cboFormatoDANFe 
            Height          =   315
            Left            =   -74880
            TabIndex        =   49
            Top             =   720
            Width           =   2415
         End
         Begin VB.ComboBox cboIndicadorPagamento 
            Height          =   315
            Left            =   -74820
            TabIndex        =   36
            Top             =   660
            Width           =   3135
         End
         Begin VB.ComboBox cboModFrete 
            Height          =   315
            Left            =   -74880
            TabIndex        =   35
            Top             =   660
            Width           =   3975
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
            Height          =   3555
            Left            =   120
            TabIndex        =   99
            Top             =   300
            Width           =   12615
            Begin VB.ComboBox cboDescricao 
               Height          =   315
               Left            =   1740
               TabIndex        =   27
               Top             =   480
               Width           =   5475
            End
            Begin VB.TextBox txtCodBarra 
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtCodProduto 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6420
               TabIndex        =   106
               Top             =   240
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtCFOP 
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
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   105
               ToolTipText     =   "CFOP"
               Top             =   840
               Width           =   630
            End
            Begin VB.TextBox txtICMS 
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
               Left            =   5460
               MaxLength       =   10
               TabIndex        =   104
               ToolTipText     =   "ICMS"
               Top             =   840
               Width           =   570
            End
            Begin VB.TextBox txtQuant 
               Height          =   315
               Left            =   8340
               MaxLength       =   10
               TabIndex        =   29
               Top             =   480
               Width           =   630
            End
            Begin VB.TextBox txtSubTotal 
               Height          =   315
               Left            =   9720
               MaxLength       =   8
               TabIndex        =   31
               Top             =   480
               Width           =   1080
            End
            Begin VB.TextBox txtValor 
               Height          =   315
               Left            =   7260
               MaxLength       =   8
               TabIndex        =   28
               Top             =   480
               Width           =   1020
            End
            Begin VB.TextBox txtUnid 
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
               Left            =   1740
               MaxLength       =   10
               TabIndex        =   103
               ToolTipText     =   "Unid de Venda"
               Top             =   840
               Width           =   510
            End
            Begin VB.TextBox txtCST 
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
               Left            =   2940
               MaxLength       =   10
               TabIndex        =   102
               ToolTipText     =   "CST"
               Top             =   840
               Width           =   510
            End
            Begin VB.TextBox txtDesc 
               Height          =   315
               Left            =   9000
               MaxLength       =   10
               TabIndex        =   30
               Top             =   480
               Width           =   690
            End
            Begin VB.TextBox txtNCM 
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
               Left            =   3480
               MaxLength       =   10
               TabIndex        =   101
               ToolTipText     =   "NCM"
               Top             =   840
               Width           =   1965
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   5040
               TabIndex        =   100
               Top             =   1800
               Visible         =   0   'False
               Width           =   810
            End
            Begin MSFlexGridLib.MSFlexGrid GridNotasItens 
               Height          =   2055
               Left            =   120
               TabIndex        =   33
               Top             =   1260
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   3625
               _Version        =   393216
               Appearance      =   0
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarItem 
               Height          =   315
               Left            =   8040
               TabIndex        =   32
               Top             =   900
               Width           =   1335
               _ExtentX        =   2355
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
               MICON           =   "NFe_Completa.frx":126C4
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
               Left            =   9420
               TabIndex        =   34
               Top             =   900
               Width           =   1335
               _ExtentX        =   2355
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
               MICON           =   "NFe_Completa.frx":126E0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               Height          =   195
               Left            =   1740
               TabIndex        =   116
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblCodFabrica 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. de Barra"
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal"
               Height          =   195
               Left            =   9720
               TabIndex        =   114
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   7260
               TabIndex        =   113
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde"
               Height          =   195
               Left            =   8340
               TabIndex        =   112
               Top             =   240
               Width           =   345
            End
            Begin VB.Label lblValorNota 
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
               Left            =   9420
               TabIndex        =   111
               Top             =   3300
               Width           =   225
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desc."
               Height          =   195
               Left            =   9000
               TabIndex        =   110
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblSubTotal 
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
               Left            =   6840
               TabIndex        =   109
               Top             =   3300
               Width           =   225
            End
            Begin VB.Label lblTotalDesc 
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
               Left            =   8580
               TabIndex        =   108
               Top             =   3300
               Width           =   225
            End
            Begin VB.Label lblTipoConsulta 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "0"
               Height          =   195
               Left            =   4200
               TabIndex        =   107
               Top             =   240
               Visible         =   0   'False
               Width           =   90
            End
         End
         Begin TabDlg.SSTab SSTab4 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   130
            Top             =   360
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   5953
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
            TabPicture(0)   =   "NFe_Completa.frx":126FC
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "txtInfComple"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Informações Adicionais"
            TabPicture(1)   =   "NFe_Completa.frx":12718
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "txtInfAdicionais"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.TextBox txtInfAdicionais 
               Height          =   2865
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   48
               Top             =   420
               Width           =   12360
            End
            Begin VB.TextBox txtInfComple 
               Height          =   2745
               Left            =   -74880
               MultiLine       =   -1  'True
               TabIndex        =   47
               Text            =   "NFe_Completa.frx":12734
               Top             =   420
               Width           =   12360
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   131
            Top             =   1140
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   4048
            _Version        =   393216
            Tabs            =   4
            Tab             =   1
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
            TabPicture(0)   =   "NFe_Completa.frx":12798
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame6"
            Tab(0).Control(1)=   "cboTransporte"
            Tab(0).Control(2)=   "txtCodTransporte"
            Tab(0).Control(3)=   "Label7"
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Volumes"
            TabPicture(1)   =   "NFe_Completa.frx":127B4
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label13"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label12"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label10"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label17"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "Label18"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "Label11"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "txtVolPesoLiquido"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "txtVolNumeracao"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "txtVolMarca"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "txtVolEspecie"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "txtVolQuant"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "txtVolPesoBruto"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).ControlCount=   12
            TabCaption(2)   =   "Reboques / Outros"
            TabPicture(2)   =   "NFe_Completa.frx":127D0
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame1"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Retenção do ICMS"
            TabPicture(3)   =   "NFe_Completa.frx":127EC
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            Begin VB.Frame Frame1 
               Caption         =   "Identificação"
               Height          =   1095
               Left            =   -74880
               TabIndex        =   147
               Top             =   420
               Width           =   12375
               Begin VB.TextBox txtPlacaReboque 
                  Height          =   315
                  Left            =   180
                  MaxLength       =   8
                  TabIndex        =   150
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.TextBox txtUFReboque 
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   2
                  TabIndex        =   149
                  Top             =   600
                  Width           =   465
               End
               Begin VB.TextBox txtRNTCReboque 
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   2
                  TabIndex        =   148
                  Top             =   600
                  Width           =   7305
               End
               Begin VB.Label Label49 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Placa"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   153
                  Top             =   360
                  Width           =   405
               End
               Begin VB.Label Label48 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   1500
                  TabIndex        =   152
                  Top             =   360
                  Width           =   210
               End
               Begin VB.Label Label47 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RNTC"
                  Height          =   195
                  Left            =   2040
                  TabIndex        =   151
                  Top             =   360
                  Width           =   450
               End
            End
            Begin VB.TextBox txtVolPesoBruto 
               Height          =   315
               Left            =   4980
               MaxLength       =   50
               TabIndex        =   146
               Top             =   780
               Width           =   1545
            End
            Begin VB.TextBox txtVolQuant 
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   145
               Top             =   780
               Width           =   825
            End
            Begin VB.TextBox txtVolEspecie 
               Height          =   315
               Left            =   960
               MaxLength       =   50
               TabIndex        =   144
               Top             =   780
               Width           =   1245
            End
            Begin VB.TextBox txtVolMarca 
               Height          =   315
               Left            =   2220
               MaxLength       =   50
               TabIndex        =   143
               Top             =   780
               Width           =   1665
            End
            Begin VB.TextBox txtVolNumeracao 
               Height          =   315
               Left            =   3900
               MaxLength       =   50
               TabIndex        =   142
               Top             =   780
               Width           =   1065
            End
            Begin VB.TextBox txtVolPesoLiquido 
               Height          =   315
               Left            =   6540
               MaxLength       =   50
               TabIndex        =   141
               Top             =   780
               Width           =   1545
            End
            Begin VB.Frame Frame6 
               Caption         =   "Veículo"
               Height          =   1095
               Left            =   -74880
               TabIndex        =   134
               Top             =   1080
               Width           =   12375
               Begin VB.TextBox txtTransRNTC 
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   2
                  TabIndex        =   137
                  Top             =   600
                  Width           =   7305
               End
               Begin VB.TextBox txtPlacaUF 
                  Height          =   315
                  Left            =   1500
                  MaxLength       =   2
                  TabIndex        =   136
                  Top             =   600
                  Width           =   465
               End
               Begin VB.TextBox txtPlaca 
                  Height          =   315
                  Left            =   180
                  MaxLength       =   8
                  TabIndex        =   135
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label Label46 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RNTC"
                  Height          =   195
                  Left            =   2040
                  TabIndex        =   140
                  Top             =   360
                  Width           =   450
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   1500
                  TabIndex        =   139
                  Top             =   360
                  Width           =   210
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Placa"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   138
                  Top             =   360
                  Width           =   405
               End
            End
            Begin VB.ComboBox cboTransporte 
               Height          =   315
               Left            =   -74880
               TabIndex        =   133
               Top             =   660
               Width           =   7695
            End
            Begin VB.TextBox txtCodTransporte 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   -67800
               MaxLength       =   50
               TabIndex        =   132
               Top             =   360
               Width           =   600
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Espécie"
               Height          =   195
               Left            =   960
               TabIndex        =   160
               Top             =   540
               Width           =   570
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso liquido"
               Height          =   195
               Left            =   6540
               TabIndex        =   159
               Top             =   540
               Width           =   855
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso bruto"
               Height          =   195
               Left            =   4980
               TabIndex        =   158
               Top             =   540
               Width           =   765
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde Vol."
               Height          =   195
               Left            =   120
               TabIndex        =   157
               Top             =   540
               Width           =   660
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Marca"
               Height          =   195
               Left            =   2220
               TabIndex        =   156
               Top             =   540
               Width           =   450
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Numeração"
               Height          =   195
               Left            =   3900
               TabIndex        =   155
               Top             =   540
               Width           =   825
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transportadora"
               Height          =   195
               Left            =   -74880
               TabIndex        =   154
               Top             =   420
               Width           =   1080
            End
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave"
            Height          =   195
            Left            =   -74820
            TabIndex        =   203
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave"
            Height          =   195
            Left            =   -74760
            TabIndex        =   165
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formato Impressão DANFE"
            Height          =   195
            Left            =   -74880
            TabIndex        =   164
            Top             =   480
            Width           =   1920
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Emissão NFe"
            Height          =   195
            Left            =   -72420
            TabIndex        =   163
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador Forma Pagto"
            Height          =   195
            Left            =   -74820
            TabIndex        =   162
            Top             =   420
            Width           =   1605
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade do Frete"
            Height          =   195
            Left            =   -74880
            TabIndex        =   161
            Top             =   420
            Width           =   1455
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   13440
         TabIndex        =   53
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
         MICON           =   "NFe_Completa.frx":12808
         PICN            =   "NFe_Completa.frx":12824
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
         Left            =   13440
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
         MICON           =   "NFe_Completa.frx":145B6
         PICN            =   "NFe_Completa.frx":145D2
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
         Left            =   13440
         TabIndex        =   55
         Top             =   3060
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":16364
         PICN            =   "NFe_Completa.frx":16380
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
         Left            =   13440
         TabIndex        =   52
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
         MICON           =   "NFe_Completa.frx":18112
         PICN            =   "NFe_Completa.frx":1812E
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
         Left            =   13440
         TabIndex        =   54
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":19EC0
         PICN            =   "NFe_Completa.frx":19EDC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   13440
         TabIndex        =   61
         Top             =   7020
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
         MICON           =   "NFe_Completa.frx":1BC6E
         PICN            =   "NFe_Completa.frx":1BC8A
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
         Left            =   13440
         TabIndex        =   57
         Top             =   4380
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":1DA1C
         PICN            =   "NFe_Completa.frx":1DA38
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
         Height          =   615
         Left            =   13440
         TabIndex        =   56
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":1F7CA
         PICN            =   "NFe_Completa.frx":1F7E6
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
         Height          =   615
         Left            =   13440
         TabIndex        =   58
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":21578
         PICN            =   "NFe_Completa.frx":21594
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdRecalcularNF 
         Height          =   615
         Left            =   13440
         TabIndex        =   60
         Top             =   6360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Recalcular"
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
         MICON           =   "NFe_Completa.frx":21BD5
         PICN            =   "NFe_Completa.frx":21BF1
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
         Height          =   615
         Left            =   13440
         TabIndex        =   59
         Top             =   5700
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "NFe_Completa.frx":23983
         PICN            =   "NFe_Completa.frx":2399F
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
         Height          =   7395
         Left            =   -74880
         TabIndex        =   202
         Top             =   420
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   13044
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Observação"
         Height          =   195
         Left            =   13680
         TabIndex        =   180
         Top             =   8040
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15465
      TabIndex        =   63
      Top             =   0
      Width           =   15495
      Begin VB.TextBox txtCodPedido 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   14160
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Cód do Pedido"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCodNota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   12900
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "Cód da Nota"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   540
         Picture         =   "NFe_Completa.frx":24279
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NFe - Nota Fiscal Eletrônica (Completa)"
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
         Left            =   1755
         TabIndex        =   66
         Top             =   180
         Width           =   5880
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   62
      Top             =   10320
      Width           =   15510
      _ExtentX        =   27358
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20876
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
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
            TextSave        =   "00:13"
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
Private iRow As Long, iCol As Long
Private Sub AtualizarGrid_Itens()
  Dim i As Integer
   Dim sSQL As String
   
   For i = 1 To GridNotasItens.Rows - 1
      If GridNotasItens.TextMatrix(i, 1) <> "" Then
         dbData.Execute "UPDATE NotaFiscalItens SET CFOP = " & GridNotasItens.TextMatrix(i, 5) & ", CST = '" & GridNotasItens.TextMatrix(i, 4) & "', NCM = '" & GridNotasItens.TextMatrix(i, 3) & "'  WHERE CodigoNota = " & txtCodNota.Text & " AND ITEM = " & GridNotasItens.TextMatrix(i, 11) & ""
         dbData.Execute "UPDATE Produtos SET NCM = '" & GridNotasItens.TextMatrix(i, 3) & "'  WHERE CODIGO = " & GridNotasItens.TextMatrix(i, 1) & ""
      End If
   Next
End Sub

Private Sub GravarPedido()
flag = False

'On Error GoTo Err_Grava

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

'If txtCodPedido = "" Then Exit Sub

'preencher objetos da nota com o pedido
sSQL = "SELECT pedidos.*, cliente.codigo, cliente.nome as VarNome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE pedidos.cod_pedido = " & txtCodPedido & ";"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.BOF Then Mostrar_Pedido r

If r.State <> 0 Then r.Close
Set r = Nothing


If TxtCodCliente.Text = "" Then MsgBox "O campo CLIENTE é obrigatório.", vbCritical, "Online Commerce": TxtCodCliente.SetFocus: Exit Sub
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
cmdAlterar.Enabled = True
cmdExcluir.Enabled = False
'Command5.Enabled = True
cmdTransmitir.Enabled = True

'Clear_Controls
LimparCamposItens

End Sub



Sub Load_Data_Itens()
Dim sSQL As String, seq As Integer

If txtCodNota.Text = "" Then Exit Sub

    sSQL = "SELECT MAX(Item) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    seq = SQLExecutaRetorno(sSQL, "r", 0) + 1
    Tb("CodigoNota") = Format(txtCodNota.Text, "@")
    Tb("Item") = seq
    Tb("CodigoProduto") = Format(txtCodProduto, "@")
    Tb("NomeProduto") = UCase(Format(cboDescricao, "@"))
    Tb("CFOP") = Format(txtCFOP, "@")
    Tb("NCM") = Format(txtNCM, "@")
    Tb("CST") = Right(Format(txtCST, "@"), 3)
    Tb("UnidadeComercial") = UCase(Format(txtUnid, "@"))
    Tb("ValorUnitarioComercializacao") = CDbl(Format(txtValor, "@"))
    Tb("ValorTotalBruto") = CDbl(Format(txtSubTotal, "@"))
    Tb("tipodesconto") = Format(1, "@")
    Tb("desconto") = CDbl(Format(txtDesc, "@"))
    Tb("Valordesconto") = CDbl(Format(txtDesc, "@"))
    If txtQuant.Text <> "" Then Tb("QuantidadeComercial") = CDbl(Format(txtQuant, "@")) Else Tb("QuantidadeComercial") = Format(1, "@")
    If txtICMS.Text <> "" Then Tb("pICMS") = CDbl(Format(txtICMS, "@"))
    If txtICMS.Text <> "" Then Tb("vBC") = CDbl(Format(txtSubTotal, "@"))
    If txtICMS.Text <> "" Then Tb("vICMS") = Round(Format(txtSubTotal, "@") * (Format(txtICMS, "@") / 100), 2)
End Sub

Private Sub Calcular_Total()
Dim var_Quant As Double
Dim var_VALOR As Currency, var_Total As Currency

If txtQuant.Text = "" Then var_Quant = 1 Else var_Quant = txtQuant.Text
If txtValor.Text = "" Then var_VALOR = 0 Else var_VALOR = txtValor.Text

var_Total = var_VALOR * var_Quant
txtSubTotal.Text = Format(var_Total, ocMONEY)
End Sub
Sub Clear_Controls()
Limpa_Tudo Me
mskEmissao.Mask = ""
mskSaida.Mask = ""
mskHora.Mask = ""
mskEmissao.Text = ""
mskSaida.Text = ""
mskHora.Text = ""
End Sub

Private Sub LimparCamposItens()
txtCodBarra.Text = ""
cboDescricao.Text = ""
txtCodProduto.Text = ""
txtCST = Format("0", "@")
txtNCM = Format("0", "@")
txtValor = Format("0", "@")
txtSubTotal = Format("0", "@")
txtQuant = Format("1", "@")
txtICMS = Format("0", "@")
txtCFOP = Format("0", "@")
txtDesc = Format("0", "@")
'chkDesc.Value = 0
'chkDesc.Value = 1
End Sub

Sub Load_Data()
'On Error GoTo erro
    vgDb.BeginTrans
    If txtVolPesoBruto.Text = "" Then txtVolPesoBruto.Text = "0"
    If txtVolPesoLiquido.Text = "" Then txtVolPesoLiquido.Text = "0"
    TbNotas("CodigoNatureza") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("CodigoNota") = Format(txtCodNota.Text, "@")
    'TbNotas("SerieNF") = 1
    'TbNotas("InscricaoEstadual") = 0
    'TbNotas("Suframa") = 0
    'TbNotas("CNPJ_CPF") = 0
    'TbNotas("Logradouro") = 0
    'TbNotas("numero") = 0
    'TbNotas("CodigoIBGE") = 0
    'TbNotas("Bairro") = 0
    'TbNotas("Complemento") = 0
    'TbNotas("Municipio") = 0
    'TbNotas("UF") = 0
    'Left$("Programar com Visual Basic é fácil", 9)
    TbNotas("InformacoesComplementares") = Format(txtInfComple, "@")
    TbNotas("NaturezaOperacao") = Format(Left(txtNatureza, 59), "@")
    TbNotas("TipoDocumento") = 1
    TbNotas("DataEmissao") = IIf(TbNotas("DataEmissao") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskEmissao, "@"))
    TbNotas("DataSaida") = IIf(TbNotas("DataSaida") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskSaida, "@"))
    TbNotas("HoraSaida") = IIf(TbNotas("HoraSaida") = Empty, Format(Time(), "HH:MM:ss"), Format(mskHora, "@"))
'    TbNotas("DataEmissao") = Format(mskEmissao, "@")
'    TbNotas("DataSaida") = Format(mskSaida, "@")
'    TbNotas("HoraSaida") = Format(mskHora, "@")
    
    
    TbNotas("TranspCodigo") = IIf(IsNull(Format(txtCodTransporte, "@")) Or Vazio(Format(txtCodTransporte, "@")), 0, Format(txtCodTransporte, "@"))
    TbNotas("TranspNome") = Format(cboTransporte, "@")
    TbNotas("TranspPlaca") = Format(txtPlaca, "@")
    TbNotas("VolumeQuantidade") = Format(txtVolQuant, "@")
    TbNotas("VolumeEspecie") = Format(Text11, "@")
    TbNotas("VolumeMarca") = Format(txtVolMarca, "@")
    TbNotas("VolumeNumeracao") = Format(Text13, "@")
    TbNotas("CodigoObservacao") = IIf(IsNull(Format(txtCodObservacao, "@")) Or Vazio(Format(txtCodObservacao, "@")), 0, Format(txtCodObservacao, "@"))
    TbNotas("ValorFrete") = IIf(Vazio(txtValorFrete), 0, CDbl(Format(txtValorFrete, "##0.00")))
    TbNotas("ValorSeguro") = IIf(IsNull(Format(txtValorSeguro, "@")) Or Vazio(Format(txtValorSeguro, "@")), 0, CDbl(Format(txtValorSeguro, "##0.00")))
    TbNotas("ValorOutrasDespesas") = IIf(IsNull(Format(txtValorOutrasDespesas, "@")) Or Vazio(Format(txtValorOutrasDespesas, "@")), 0, CDbl(Format(txtValorOutrasDespesas, "##0.00")))
    TbNotas("NumeroNota") = Format(txtNumNota, "@")
    TbNotas("cCodigoNota") = IIf(TbNotas("cCodigoNota") = 0, GeraCodigoNota, TbNotas("cCodigoNota"))
    TbNotas("VolumePesoBruto") = IIf(IsNull(Format(txtVolPesoBruto, "@")) Or Vazio(Format(txtVolPesoBruto, "@")), 0, CDbl(Format(txtVolPesoBruto, "##0.000")))
    TbNotas("VolumePesoLiquido") = IIf(IsNull(Format(txtVolPesoLiquido, "@")) Or Vazio(Format(txtVolPesoLiquido, "@")), 0, CDbl(Format(txtVolPesoLiquido, "##0.000")))
    TbNotas("BaseICMS") = IIf(IsNull(Format(txtBaseICMS, "@")) Or Vazio(Format(txtBaseICMS, "@")), 0, CDbl(Format(txtBaseICMS, "##0.000")))
    TbNotas("BaseICMSST") = IIf(IsNull(Format(txtBaseICMSST, "@")) Or Vazio(Format(txtBaseICMSST, "@")), 0, CDbl(Format(txtBaseICMSST, "##0.00")))
    TbNotas("TranspPlacaUF") = Format(txtPlacaUF, "@")
    TbNotas("CodigoCorrentista") = IIf(IsNull(Format(TxtCodCliente, "@")) Or Vazio(Format(TxtCodCliente, "@")), 0, Format(TxtCodCliente, "@"))
    TbNotas("RazaoSocial") = Format(CboCliente, "@")
    TbNotas("ModFrete") = IIf(IsNull(Format(Left(cboModFrete.Text, 1), "@")) Or Vazio(Format(Left(cboModFrete.Text, 1), "@")), 9, Format(Left(cboModFrete.Text, 1), "@"))
    TbNotas("IndicadorFormaPagamento") = Format(cboIndicadorPagamento.Text, "@")
    TbNotas("FormatoImpressaoDANFE") = Format(cboFormatoDANFe.Text, "@")
    TbNotas("FormatoEmissaoNFe") = Format(cboTipoEmissao.Text, "@")
    TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")
    TbNotas("IdentificadorDestino") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("IndicadorIEDestinatario") = IIf(IsNull(Format(Left(cboTipoContribuinte.Text, 1), "@")) Or Vazio(Format(Left(cboTipoContribuinte.Text, 1), "@")), 1, Format(Left(cboTipoContribuinte.Text, 1), "@"))
    TbNotas("ConsumidorFinal") = IIf(IsNull(Format(Left(cboConsumidorFinal.Text, 1), "@")) Or Vazio(Format(Left(cboConsumidorFinal.Text, 1), "@")), 1, Format(Left(cboConsumidorFinal.Text, 1), "@"))
   
    TbNotas("ChavedeAcessoAdicional") = Format(txtChaveReferenciada.Text, "@")
   
    'cboConsumidorFinal
    Exit Sub

Resume

'erro:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
'    vgDb.RollbackTrans
'    Exit Sub
End Sub
Public Sub Load_Controls()
'On Error GoTo erro
    TxtCodCliente = Format(TbNotas("CodigoCorrentista"), "@")
    CboCliente = Format(TbNotas("RazaoSocial"), "@")
    
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
    'cboDestOperacao = Format(TbNotas("CodigoNatureza"), "@")
    
    txtInfAdicionais = Format(TbNotas("InformacoesAdicionais"), "@")
    txtNatureza = Format(TbNotas("NaturezaOperacao"), "@")
    mskEmissao = Format(TbNotas("DataEmissao"), "@")
    mskSaida = Format(TbNotas("DataSaida"), "@")
    mskHora = Format(TbNotas("HoraSaida"), "@")
    txtCodTransporte = Format(TbNotas("TranspCodigo"), "@")
    cboTransporte = Format(TbNotas("TranspNome"), "@")
    txtPlaca = Format(TbNotas("TranspPlaca"), "@")
    txtVolQuant = Format(TbNotas("VolumeQuantidade"), "@")
    Text11 = Format(TbNotas("VolumeEspecie"), "@")
    txtVolMarca = Format(TbNotas("VolumeMarca"), "@")
    Text13 = Format(TbNotas("VolumeNumeracao"), "@")
    txtCodObservacao = Format(TbNotas("CodigoObservacao"), "@")
    txtNumNota = Format(TbNotas("NumeroNota"), "@")
    txtValorSeguro = Format(TbNotas("ValorSeguro"), "##,##0.00")
    txtValorOutrasDespesas = Format(TbNotas("ValorOutrasDespesas"), "##,##0.00")
    txtValorFrete = Format(TbNotas("ValorFrete"), "##,##0.00")
    txtBaseICMS = Format(TbNotas("BaseICMS"), "##,##0.00")
    txtBaseICMSST = Format(TbNotas("BaseICMSST"), "##,##0.00")
    txtVolPesoBruto = Format(TbNotas("VolumePesoBruto"), "@")
    txtVolPesoLiquido = Format(TbNotas("VolumePesoLiquido"), "@")
    txtPlacaUF = Format(TbNotas("TranspPlacaUF"), "@")

    If TbNotas("ModFrete") = 0 Then
        cboModFrete.Text = "0 - POR CONTA DO EMITENTE"
    ElseIf TbNotas("ModFrete") = 1 Then
        cboModFrete.Text = "1 - POR CONTA DE AMBOS"
    ElseIf TbNotas("ModFrete") = 2 Then
        cboModFrete.Text = "2 - POR CONTA DO TERCEIROS"
    ElseIf TbNotas("ModFrete") = 9 Then
        cboModFrete.Text = "9 - SEM FRETE"
    End If
    
    txtCodNota = Format(TbNotas("CodigoNota"), "@")
    Text30 = Format(TbNotas("ChavedeAcesso"), "@")
    Text31 = Format(TbNotas("NumeroProtocolo"), "@")
    Text32 = Format(TbNotas("NumeroRecibo"), "@")
    cboIndicadorPagamento.Text = Format(TbNotas("IndicadorFormaPagamento"), "@")
    cboFormatoDANFe.Text = Format(TbNotas("FormatoImpressaoDANFE"), "@")
    cboTipoEmissao.Text = Format(TbNotas("FormatoEmissaoNFe"), "@")
    txtCodPedido = Format(TbNotas("cod_pedido"), "@")
    txtInfComple.Text = Format(TbNotas("InformacoesComplementares"), "@")
    
    txtChaveReferenciada.Text = Format(TbNotas("ChavedeAcessoAdicional"), "@")

    MostraStatus = MostraStatus_F9()
    
    frmNota.Enabled = True
    'frmTransmissao.Enabled = True
    frmItens.Enabled = True
Exit Sub

Resume

'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub Mostrar_ItensNota()
Dim sSQL As String, enviada As Boolean
Dim totalRegistros As Long
    
    'On Error GoTo ErrLoad
    
    sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL
    
    If Tb.RecordCount > 0 Then totalRegistros = Tb.RecordCount
    
    
'    enviada = SQLExecutaRetorno("SELECT Enviada FROM NotaFiscal WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text), "Enviada", 0)
'
'    If enviada Then
'       cboDestOperacao.Enabled = False
'       Text3.Enabled = False
'       mskEmissao.Enabled = False
'       mskSaida.Enabled = False
'       mskHora.Enabled = False
'       Text7.Enabled = False
'       Text8.Enabled = False
'       txtPlaca.Enabled = False
'       txtVolQuant.Enabled = False
'    End If
    
    LimparGridItensNota
    DoEvents
    FormatarGridItensNota Tb
    Exit Sub
    
'ErrLoad:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
End Sub

Private Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 12
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 0
      .ColWidth(2) = 3500
      .ColWidth(3) = 1200
      .ColWidth(4) = 500
      .ColWidth(5) = 800
      .ColWidth(6) = 500
      .ColWidth(7) = 1000
      .ColWidth(8) = 800
      .ColWidth(9) = 800
      .ColWidth(10) = 1000
      .ColWidth(11) = 0

      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓD."
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "NCM"
      .TextMatrix(0, 4) = "CST"
      .TextMatrix(0, 5) = "CFOP"
      .TextMatrix(0, 6) = "UND"
      .TextMatrix(0, 7) = "VALOR"
      .TextMatrix(0, 8) = "QTDE"
      .TextMatrix(0, 9) = "DESC."
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = "ITEM"
      
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
      .ColAlignment(7) = 2
      .ColAlignment(8) = 2
      .ColAlignment(9) = 2
      .ColAlignment(10) = 2
      .ColAlignment(11) = 1
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("CodigoProduto")
            .TextMatrix(.Rows - 1, 2) = rTabela("NomeProduto")
            .TextMatrix(.Rows - 1, 3) = rTabela("NCM")
            .TextMatrix(.Rows - 1, 4) = rTabela("CST")
            .TextMatrix(.Rows - 1, 5) = rTabela("CFOP")
            .TextMatrix(.Rows - 1, 6) = rTabela("UnidadeComercial")
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("ValorUnitarioComercializacao"), ocMONEY)
            .TextMatrix(.Rows - 1, 8) = Format(rTabela("QuantidadeComercial"), ocMONEY)
            .TextMatrix(.Rows - 1, 9) = Format(rTabela("valordesconto"), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = Format(rTabela("ValorTotalBruto"), ocMONEY)
            .TextMatrix(.Rows - 1, 11) = rTabela("ITEM")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
              
     'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
      GridNotasItens.Col = 0
            
      .Visible = True
      .Redraw = True
      
      SomarGridItens
   End With
End Sub

Private Sub PreencherGridNotas()
Dim totalRegistros As Long

On Error GoTo ErrLoad

RsOpen TbConsulta, "SELECT *, " & _
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
Dim sSQL As String, vTotal As Double

On Error GoTo erro

    sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    
    Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub TransformarPedidoemNFE()
Dim sSQL As String
'Dim r As ADODB.Recordset
'Dim totalRegistros As Long

'If txtCodPedido = "" Then Exit Sub

'preencher objetos da nota com o pedido
'sSQL = "SELECT pedidos.*, cliente.codigo, cliente.nome as VarNome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE pedidos.cod_pedido = " & txtCodPedido & ";"
'Set r = dbData.OpenRecordset(sSQL, totalRegistros)

'If Not r.BOF Then Mostrar_Pedido r

'If r.State <> 0 Then r.Close
'Set r = Nothing

Dim tblItensPedido As ADODB.Recordset

'Atualiza a base de dados (funcionando)
Dim VarCodNota As Integer
VarCodNota = CInt(txtCodNota.Text)

sSQL = "INSERT INTO NotaFiscalItens ( " & _
        "CodigoProduto, " & _
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
        "SELECT pedidos_itens.cod_produto, produtos.descricao, produtos.cfop, produtos.ncm, produtos.icmscst, produtos.unid_medida, pedidos_itens.preco, (pedidos_itens.preco * pedidos_itens.quantidade) as varValorBruto, 1, 0, 0, pedidos_itens.quantidade, 0, (pedidos_itens.preco * pedidos_itens.quantidade) as varVBC, 0, pedidos_itens.item, " & VarCodNota & " " & _
        "FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
        "WHERE pedidos_itens.COD_PEDIDO = " & txtCodPedido.Text & ";"
'MsgBox sSQL
dbData.Execute sSQL


'preencher o grid dos itens com o pedido
sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

FormatarGridItensNota Tb

'MOSTRAR A QUANTIDADE REGISTROS
'lblQuantPedidos.Caption = Format(totalRegistros, "00")
End Sub

Private Sub cboAno_Change()

End Sub

Private Sub cboAno_GotFocus()
SelectControl mskFim
End Sub


Private Sub cboCategoria_Change()

End Sub

Private Sub cboCategoria_GotFocus()
SelectControl mskFim
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
Dim sSQL As String
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
Dim sSQL As String
Dim r As ADODB.Recordset
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

moCombo.AttachTo cboConsumidorFinal
End Sub


Private Sub cboDestOperacao_GotFocus()
Dim VarText As String
VarText = cboDestOperacao.Text

cboDestOperacao.Clear
cboDestOperacao.AddItem "1 - Operação Interna"
cboDestOperacao.AddItem "2 - Operação Interestadual"
cboDestOperacao.AddItem "3 - Operação com Exterior"

If cboDestOperacao.Text = "" Then cboDestOperacao.Text = VarText

moCombo.AttachTo cboDestOperacao
End Sub


Private Sub cboFiltroNota_Click()
lblConNotaNumNota.Visible = False
txtConNotaNumNota.Visible = False
lblConNotaCliente.Visible = False
cboConNotaCliente.Visible = False
txtConNotaCodCliente.Visible = False
lblConNotaAno.Visible = False
lblConNotaMes.Visible = False
cboConNotaAno.Visible = False
cboConNotaMes.Visible = False
lblConNotaInicial.Visible = False
lblConNotaFinal.Visible = False
mskConNotaInicial.Visible = False
mskConNotaFinal.Visible = False
cmdConNotaCal1.Visible = False
cmdConNotaCal2.Visible = False

If cboFiltroNota.Text = "TODAS" Then
    Exit Sub
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    lblConNotaNumNota.Visible = True
    txtConNotaNumNota.Visible = True
ElseIf cboFiltroNota.Text = "CLIENTE" Then
    lblConNotaCliente.Visible = True
    cboConNotaCliente.Visible = True
    'txtConNotaCodCliente.Visible = True
ElseIf cboFiltroNota.Text = "DATAS" Then
    lblConNotaInicial.Visible = True
    lblConNotaFinal.Visible = True
    mskConNotaInicial.Visible = True
    mskConNotaFinal.Visible = True
    cmdConNotaCal1.Visible = True
    cmdConNotaCal2.Visible = True
ElseIf cboFiltroNota.Text = "MENSAL" Then
    lblConNotaAno.Visible = True
    lblConNotaMes.Visible = True
    cboConNotaAno.Visible = True
    cboConNotaMes.Visible = True
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

moCombo.AttachTo cboFinalidade
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


Private Sub cboMes_Change()

End Sub

Private Sub cboMes_GotFocus()
SelectControl mskFim
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
cboModFrete.AddItem "0 - POR CONTA DO EMITENTE"
cboModFrete.AddItem "1 - POR CONTA DE AMBOS"
cboModFrete.AddItem "2 - POR CONTA DO TERCEIROS"
cboModFrete.AddItem "9 - SEM FRETE"

If cboModFrete.Text = "" Then cboModFrete.Text = VarText

'cboModFrete.AddItem ""
moCombo.AttachTo cboModFrete
End Sub


Private Sub cboNatureza_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset

If cboNatureza.Text = "" Then Exit Sub

sSQL = "SELECT CodigoNatureza, NomeNatureza FROM NaturezaOperacaoNF where CodigoNatureza = " & cboNatureza.Text & "ORDER BY CodigoNatureza;"
Set r = dbData.OpenRecordset(sSQL)

txtNatureza.Text = UCase(ValidateNull(r("NomeNatureza")))

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cboTipoContribuinte_GotFocus()
Dim VarText As String
VarText = cboTipoContribuinte.Text

cboTipoContribuinte.Clear
cboTipoContribuinte.AddItem "1 - CONTRIBUINTE ICMS"
cboTipoContribuinte.AddItem "2 - CONTRIBUINTE ISENTO"
cboTipoContribuinte.AddItem "9 - NÃO CONTRIBUINTE"

If cboTipoContribuinte.Text = "" Then cboTipoContribuinte.Text = VarText

moCombo.AttachTo cboTipoContribuinte
End Sub

Private Sub cboTipoDest_GotFocus()
Dim VarText As String
VarText = cboTipoDest.Text

cboTipoDest.Clear
cboTipoDest.AddItem "CLIENTE"
cboTipoDest.AddItem "FORNECEDOR"

If cboTipoDest.Text = "" Then cboTipoDest.Text = VarText

moCombo.AttachTo cboTipoDest
End Sub


Private Sub cboTipoNota_GotFocus()
Dim VarText As String
VarText = cboTipoNota.Text

cboTipoNota.Clear
cboTipoNota.AddItem "0 - ENTRADA"
cboTipoNota.AddItem "1 - SAÍDA"

If cboTipoNota.Text = "" Then cboTipoNota.Text = VarText

moCombo.AttachTo cboTipoNota
End Sub


Private Sub cmdAdicionarItem_Click()
Dim sSQL As String, vTotal As Double

'On Error GoTo erro
If txtCodNota.Text = "" Then Exit Sub
If txtCodProduto.Text = "" Then Exit Sub
If txtSubTotal.Text = "" Then Exit Sub
If Len(txtNCM.Text) < 8 Then ShowMsg "NCM INCORRETO!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbExclamation

    sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL


    vgDb.BeginTrans
    
    'insere os dados itens
    Tb.AddNew
    Load_Data_Itens
    Tb.Update
    
    vgDb.CommitTrans
    
    Call cmdRecalcularNF_Click
    
    'Limpa_Tudo Me ' limpa tudo
    
    'sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    'vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    'sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    'SQLExecuta sSQL
    
    'EXIBIR NO GRID
    sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL
    
    FormatarGridItensNota Tb
    
    LimparCamposItens
    
    KeyCode = 0
    lblTipoConsulta.Caption = "0"
    txtCodBarra.SetFocus
'End If
Exit Sub
'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub
Private Sub SomarGridItens()
Dim Total As Currency, SUBTOTAL As Currency, Desc As Currency
Dim i As Integer

SUBTOTAL = 0
Desc = 0
Total = 0

'Sub-Total
With GridNotasItens
   For i = 1 To .Rows - 1
      .Col = 0
      .Row = i
      
         .Col = 6
         SUBTOTAL = SUBTOTAL + (.TextMatrix(.Row, 7) * .TextMatrix(.Row, 8))
        '.Col = 10
         Desc = Desc + .TextMatrix(.Row, 9)
         '.Col = 11
         Total = Total + .TextMatrix(.Row, 10)
   Next
End With


lblSubTotal.Caption = Format(SUBTOTAL, ocMONEY)
lblTotalDesc.Caption = Format(Desc, ocMONEY)
lblValorNota.Caption = Format(Total, ocMONEY)
'txtTotaldaNota
End Sub

Private Sub cmdAlterar_Click()
Dim sSQL As String
'Dim enviada As Boolean
'Dim totalRegistros As Long
flag = False

'On Error GoTo Err_Grava

If TxtCodCliente.Text = "" Then MsgBox "O campo código do cliente é obrigatório.", vbCritical, "Online Commerce": TxtCodCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo código CFOP é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
If txtCodNota.Text = "" Then MsgBox "Essa Nota Fiscal não existe!", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
    
If TbNotas.EditMode = 2 Then
   resp = MsgBox("Confirma inclusão ?", 36, Titulo)
   flag = True
   If resp <> 6 Then Exit Sub
Else
   resp = MsgBox("Confirma alteração ?", 36, Titulo)
   flag = False
   If resp <> 6 Then Exit Sub
   
   'TbNotas.Edit
End If

Load_Data
TbNotas.Update

vgDb.CommitTrans

Load_Controls

SomarProdutosNota

PreencherGridNotas

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdTransmitir.Enabled = False
cmdImprimir.Enabled = False
cmdCancelarNota.Enabled = False
cmdConsultar.Enabled = False
frmNota.Enabled = False
'frmTransmissao.Enabled = False
frmItens.Enabled = False
SSTab3.Tab = 0
SSTab2.Tab = 0

Clear_Controls
LimparCamposItens

Exit Sub
Resume
'Err_Grava:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
End Sub

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
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
 If MsgBox("Tem certeza que deseja Cancelar a Nota Fiscal?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
 Justificativa = InputBox("Informe a Justificativa do Cancelamento da Nota:", "Cancelamento da Nota", "DESISTENCIA DA COMPRA")
 vsNumeroNota = Val(txtCodNota.Text)
 iRetorno = CancelaNFe(TbNotas("ChavedeAcesso"), TbNotas("NumeroProtocolo"), Justificativa, True)
 If iRetorno Then
    SQL = "UPDATE NotaFiscal SET " & _
          "Cancelada = 1, " & _
          "CanceladaProtocolo = " & NFeNumeroProtocolo & ", " & _
          "Justificativa = '" & Justificativa & "' " & _
          "WHERE CodigoNota = " & Val(txtCodNota.Text)
    vgDb.Execute SQL
    RsOpen TbNotas, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & Val(txtCodNota.Text)
    Load_Controls
    FormatarGridNotas TbNotas
 End If
End Sub

Private Sub cboCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = CboCliente.Text
codAtual = TxtCodCliente.Text
CboCliente.Clear

If cboTipoDest.Text = "FORNECEDOR" Then
    sSQL = "SELECT DISTINCT RAZAO, codigo FROM FORNECEDOR ORDER BY razao;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       CboCliente.AddItem r("RAZAO")
       CboCliente.ItemData(CboCliente.NewIndex) = r("codigo")
       r.MoveNext
    Loop
Else
    sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
    Set r = dbData.OpenRecordset(sSQL)
    
    Do While Not r.EOF
       CboCliente.AddItem r("nome")
       CboCliente.ItemData(CboCliente.NewIndex) = r("codigo")
       r.MoveNext
    Loop
End If



If r.State <> 0 Then r.Close
Set r = Nothing

CboCliente.Text = itemAtual
TxtCodCliente.Text = codAtual
moCombo.AttachTo CboCliente
End Sub

Private Sub cboCliente_LostFocus()
On Error GoTo TrataErro

If CboCliente.Text = "" Then TxtCodCliente.Text = "": Exit Sub
If CboCliente.ListIndex = -1 Then TxtCodCliente.Text = "": Exit Sub

TxtCodCliente = CboCliente.ItemData(CboCliente.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboDescricao_GotFocus()
moCombo.AttachTo cboDescricao
   
Dim sSQL As String
Dim r As ADODB.Recordset

If lblTipoConsulta.Caption = "0" Or lblTipoConsulta.Caption = "2" Then

    If cboDescricao.ListIndex = -1 Then
        cboDescricao.Clear
        
        sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboDescricao.AddItem ValidateNull(r("descricao"))
            cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
           r.MoveNext
        Loop
    End If
End If

moCombo.AttachTo cboDescricao
End Sub


Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboDescricao_LostFocus()
'If chkDesc.Value = 1 Then
If lblTipoConsulta.Caption = "0" Or lblTipoConsulta.Caption = "2" Then
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'txtCodBarra.Text = ""
   
   'If cboDescricao.Text = "" Then txtCodProduto.Text = "": Exit Sub
   'If cboDescricao.ListIndex = -1 Then txtCodProduto.Text = "": Exit Sub
   
    If cboDescricao.Text = "" Then
        txtCodProduto.Text = ""
        lblTipoConsulta.Caption = "0"
        txtCodBarra.Locked = False
        txtCodBarra.Text = ""
        txtUnid.Text = ""
        txtCFOP.Text = ""
        txtCST.Text = ""
        txtNCM.Text = ""
        txtICMS.Text = ""
        txtValor.Text = "0"
        txtSubTotal.Text = "0"
        Exit Sub
    End If
    
    If cboDescricao.ListIndex = -1 Then txtCodProduto.Text = "": lblTipoConsulta.Caption = "0": txtCodBarra.Locked = False: cboDescricao.Text = "": txtCodBarra.Text = "": Exit Sub


   'txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)
   
   'If txtCodProduto.Text = "" Then Exit Sub
   
   txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)
   
    If txtCodProduto.Text = "" Then
        lblTipoConsulta.Caption = "0"
        txtCodBarra.Locked = False
        cboDescricao.Text = ""
        txtCodBarra.Text = ""
        Exit Sub
    End If
   
   sSQL = "SELECT codigo, descricao, cod_barra, unid_medida, ncm, CFOP, ICMSCST, ICMSAliq  FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
    If Not r.BOF Then
        If txtCodBarra.Text = "" Then txtCodBarra.Text = r("cod_barra")
        txtUnid.Text = ValidateNull(r("unid_medida"))
        txtCFOP.Text = ValidateNull(r("CFOP"))
        txtCST.Text = ValidateNull(r("ICMSCST"))
        txtNCM.Text = ValidateNull(r("ncm"))
        txtICMS.Text = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
        lblTipoConsulta.Caption = "2"
        txtCodBarra.Locked = True
        MostrarValorVenda
        txtQuant.SetFocus
    ElseIf r.BOF Then
        ShowMsg "Produto não cadastrado.", vbExclamation
        lblTipoConsulta.Caption = "0"
        cboDescricao.Text = ""
        txtCodBarra.Text = ""
        txtUnid.Text = ""
        txtCFOP.Text = ""
        txtCST.Text = ""
        txtNCM.Text = ""
        txtICMS.Text = ""
        txtValor.Text = "0"
        txtSubTotal.Text = "0"
        txtCodBarra.Locked = False
        If r.State <> 0 Then r.Close
    End If

   'If r.BOF Then ShowMsg "Produto não cadastrado.", vbExclamation
   
End If
End Sub
Private Sub MostrarValorVenda()
Dim sSQL As String
Dim r As ADODB.Recordset
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
Dim sSQL As String
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
moCombo.AttachTo cboNatureza
End Sub

Private Sub cboTransporte_GotFocus()
Dim sSQL As String
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


Private Sub chkCodBarra_Click()
'If chkCodBarra.Value = 1 Then chkDesc.Value = 0
txtCodBarra.Locked = False
cboDescricao.Text = ""
If txtCodBarra.Locked = False Then cboDescricao.Locked = True
txtCodBarra.SetFocus
End Sub

Private Sub chkDesc_Click()
'If chkDesc.Value = 1 Then chkCodBarra.Value = 0
cboDescricao.Locked = False
txtCodBarra.Text = ""
If cboDescricao.Locked = False Then txtCodBarra.Locked = True
If frmItens.Enabled = True Then cboDescricao.SetFocus
End Sub


Private Sub cmdCancelar_Click()
'On Error GoTo Err_Cancela

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdTransmitir.Enabled = False
cmdImprimir.Enabled = False
cmdCancelarNota.Enabled = False
cmdConsultar.Enabled = False
frmNota.Enabled = False
'frmTransmissao.Enabled = False
frmItens.Enabled = False
lblTipoConsulta.Caption = "0"
SSTab3.Tab = 0
SSTab2.Tab = 0

If TbNotas.EditMode <> 0 Then TbNotas.CancelUpdate

Limpa_Tudo Me
mskEmissao.Mask = ""
mskSaida.Mask = ""
mskHora.Mask = ""
mskEmissao.Text = ""
mskSaida.Text = ""
mskHora.Text = ""
txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
Exit Sub

'Err_Cancela:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
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

Private Sub cmdConsultar_Click()
    If (Text32.Text = Empty) Or (Text32.Text = "") Then Exit Sub
    If (Text30.Text = "") And (Text32.Text = "0") Then Exit Sub
    If (Text32.Text <> Empty) Or (Text32.Text <> "0") Then
       vsNumeroNota = Val(txtCodNota.Text)
       ConsultaRecibo Text32.Text, Text30.Text, "1", True
    Else
       consultaNFe Text30.Text
       If cStat = 100 Then
          SQL = "UPDATE NotaFiscal SET " & _
                "Enviada = 1, " & _
                "NumeroProtocolo = " & NFeNumeroProtocolo & ", " & _
                "DataHoraProcotolo = '" & NFeDataHora & "' " & _
                "WHERE CodigoNota = " & Val(txtCodNota.Text)
          vgDb.Execute SQL
          'SQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
          '      "(" & Val(txtCodNota.Text) & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
          'vgDb.Execute SQL
       End If
    End If
    RsOpen TbNotas, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & Val(txtCodNota.Text)
    Load_Controls
    FormatarGridNotas TbNotas
End Sub

Private Sub cmdExibirConNotas_Click()
Dim totalRegistros As Long

On Error GoTo ErrLoad

If cboFiltroNota.Text = "TODAS" Then
    RsOpen TbConsulta, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    RsOpen TbConsulta, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE NumeroNota = " & txtConNotaNumNota & " order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "CLIENTE" Then
    RsOpen TbConsulta, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoCorrentista = " & txtConNotaCodCliente.Text & " order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "DATAS" Then
    RsOpen TbConsulta, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) order by NumeroNota desc"
ElseIf cboFiltroNota.Text = "MENSAL" Then
    RsOpen TbConsulta, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE (MONTH(DataEmissao) = " & cboConNotaMes.ListIndex + 1 & ") AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by NumeroNota desc"
End If

If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount
lblTotalNota.Caption = Format(totalRegistros, "00")

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub cmdNovo_Click()
'On Error GoTo ErrLoad

'pegando o numero correto da nota
'Dim var_NumeroNota As Integer
Dim ConsultaSQL As String
Dim tbNota As ADODB.Recordset

      ConsultaSQL = "SELECT ISNULL(MAX(numeronota), 0) AS Maior_nota FROM NotaFiscal"
      Set tbNota = dbData.OpenRecordset(ConsultaSQL)
      'If Not tbNota.BOF Then var_NumeroNota = tbNota("ultima_nota") + 1
      
'preecher objetos do form
Dim totalRegistros As Long
RsOpen TbNotas, "SELECT *, " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM NotaFiscal"
                
If TbNotas.RecordCount > 0 Then totalRegistros = TbNotas.RecordCount

    Clear_Controls
    
    txtCodPedido.Text = 0
    
    If TbNotas.EOF And TbNotas.BOF Then
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = "1"
    Else
        TbNotas.MoveLast
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = TbNotas("CodigoNota") + 1
    End If
    
    
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
    cboModFrete.Text = ""
    txtValorFrete.Text = "0,00"
    txtValorOutrasDespesas.Text = "0,00"
    txtVolPesoBruto.Text = "0,00"
    txtVolPesoLiquido.Text = "0,00"
    txtValorSeguro.Text = "0,00"
    txtBaseICMSST.Text = "0,00"
    txtBaseICMS.Text = "0,00"
    txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
    LimparCamposItens
    
    TbNotas.AddNew
    cmdNovo.Enabled = False
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdTransmitir.Enabled = False
    cmdImprimir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdConsultar.Enabled = False
    frmNota.Enabled = True
'    frmTransmissao.Enabled = True
    frmItens.Enabled = True
    cboTipoNota.SetFocus

Exit Sub
Resume

'ErrLoad:
'    MsgBox Err.Description, vbCritical
'    Err.Clear
'    Set TbNotas = Nothing
End Sub

Private Sub cmdRecalcularNF_Click()
Dim sSQL As String
    sSQL = "UPDATE NotaFiscal SET " & _
           "ValorProdutos = tb.TotalBruto, " & _
           "ValorDesconto = tb.Desconto, " & _
           "ValorNota = (Tb.TotalBruto - Tb.Desconto) " & _
           "FROM (SELECT ISNULL(SUM((QuantidadeComercial * ValorUnitarioComercializacao)), 0) TotalBruto, ISNULL(SUM(ValorDesconto), 0) Desconto FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ") tb " & _
           "WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End Sub

Private Sub cmdRemoverItem_Click()
Dim sSQL As String, vTotal As Double

'On Error GoTo erro

    'sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    'RsOpen Tb, sSQL


    'vgDb.BeginTrans
    
    'Tb.AddNew 'insere os dados
    'Load_Data_Itens
    'Tb.Update
    
    'vgDb.CommitTrans
    
    'Limpa_Tudo Me ' limpa tudo
    If ShowMsg("Deseja remover o item: " & GridNotasItens.TextMatrix(GridNotasItens.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    
    dbData.Execute "DELETE FROM NotaFiscalItens WHERE (CodigoProduto = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 1) & ") AND (ITEM = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 11) & ");"

    Call cmdRecalcularNF_Click
    
   ' sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    'vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    'sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)

    'SQLExecuta sSQL
    
    sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL
    
    FormatarGridItensNota Tb
    
'    lblValorNota.Caption = Format(Tb("vTotal"), ocMONEY)
    
    KeyCode = 0
    'If chkDesc.Value = 1 Then
    lblTipoConsulta.Caption = "0"
    cboDescricao.SetFocus
'End If
Exit Sub
'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
flag = False

'On Error GoTo Err_Grava

If TxtCodCliente.Text = "" Then MsgBox "O campo CLIENTE é obrigatório.", vbCritical, "Online Commerce": CboCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo código CFOP é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
'If txtCodObservacao.Text = "" Then MsgBox "O campo mensagem é obrigatório.", vbCritical, "Online Commerce": txtCodObservacao.SetFocus: Exit Sub

If txtCodPedido.Text = "0" Then

Else
    RsOpen TbNotas, "SELECT * FROM NotaFiscal"
    TbNotas.AddNew
End If

If TbNotas.EditMode = 2 Then
   resp = MsgBox("Confirma inclusão ?", 36, Titulo)
   flag = True
   If resp <> 6 Then Exit Sub
Else
   resp = MsgBox("Confirma alteração ?", 36, Titulo)
   flag = False
   If resp <> 6 Then Exit Sub
   'TbNotas.Edit
End If

Load_Data
TbNotas.Update
vgDb.CommitTrans

Load_Controls

SomarProdutosNota

PreencherGridNotas

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdTransmitir.Enabled = False
cmdImprimir.Enabled = False
cmdCancelarNota.Enabled = False
cmdConsultar.Enabled = False
frmNota.Enabled = False
'frmTransmissao.Enabled = False
frmItens.Enabled = False
SSTab3.Tab = 0
SSTab2.Tab = 0

Clear_Controls
LimparCamposItens

Exit Sub
Resume
'Err_Grava:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
End Sub
Private Sub cmdExcluir_Click()
'On Error GoTo Err_Delete

If MostraStatus.Caption <> "Transmitida/Autorizada" Or MostraStatus.Caption <> "Cancelada" Or MostraStatus.Caption <> "Denegada" Then
    resp = MsgBox("Confirma a exclusão ?", 36, Titulo)
    If resp <> 6 Then Exit Sub
    
    dbData.Execute "DELETE FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    dbData.Execute "DELETE FROM NotaFiscal WHERE  CodigoNota = " & Val(txtCodNota.Text)
    PreencherGridNotas
    Limpa_Tudo Me
    mskEmissao.Mask = ""
    mskSaida.Mask = ""
    mskHora.Mask = ""
    mskEmissao.Text = ""
    mskSaida.Text = ""
    mskHora.Text = ""
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdTransmitir.Enabled = False
    cmdImprimir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdConsultar.Enabled = False
    frmNota.Enabled = False
    'frmTransmissao.Enabled = False
    frmItens.Enabled = False
    SSTab3.Tab = 0
    SSTab2.Tab = 0
    txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
Else
    MsgBox "Não é possivel excluir uma nota Cancelada, Trasmitida ou Denegada!", vbCritical, "Online Commerce": Exit Sub
End If

Exit Sub

'Err_Delete:
'    If Err = 3021 Then
'        Clear_Controls
'        Exit Sub
'    End If
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub cmdImprimir_Click()
   On Error GoTo deuErro
     Dim sistNFe As snfe.Util
     Set sistNFe = New snfe.Util
     
     dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.path)
     xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(mskEmissao, "yyyymm") & "\" & Text30 & "-procNFe.xml"
     xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & Text30 & ".pdf"
     
     If Not Existe(xCaminhoXML) Then consultaNFe Text30, True
     
     If Not Existe(xCaminhoXML) Then Exit Sub
     
     Call sistNFe.ImpNFe(xCaminhoXML, False, "", True, xCaminhoPDF, 0)
     
     Exit Sub
deuErro:
    If InStr(1, Err.Description, "Exception") > 0 Then
       Call sistNFe.ImpNFe(xCaminhoXML, False, "", True, xCaminhoPDF, 1)
    Else
       MsgBox Err.Description, vbInformation
    End If
    Err.Clear
End Sub

Private Sub cmdTransmitir_Click()
Call cmdRecalcularNF_Click
DoEvents
iRetorno = TransmitirNFe(Val(txtCodNota.Text), "1", True)
If iRetorno Then
   SQL = "UPDATE NotaFiscal SET " & _
         "Enviada = 1 " & _
         "WHERE CodigoNota = " & Val(txtCodNota.Text)
   vgDb.Execute SQL
   RsOpen TbNotas, "SELECT *, " & _
                   "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                   "FROM NotaFiscal WHERE CodigoNota = " & Val(txtCodNota.Text)
   Load_Controls
   FormatarGridNotas TbNotas
End If
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_GotFocus()

End Sub


Private Sub Combo2_Change()

End Sub

Private Sub Combo3_Change()

End Sub

Private Sub Combo4_Change()

End Sub

Private Sub Command1_Click()
PreencherGridNotas
End Sub

Private Sub Command2_Click()
If cboIndicePedidos.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
If cboIndicePedidos.Text = "PEDIDO" Then
    If txtConCodPedido.Text = "" Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE pedidos.cod_pedido = " & txtConCodPedido & ";"
ElseIf cboIndicePedidos.Text = "CLIENTE" Then
    If txtCodClientePedidos.Text = "" Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (cliente.codigo = " & txtCodClientePedidos.Text & ") ORDER BY pedidos.cod_pedido;"
ElseIf cboIndicePedidos.Text = "DATAS" Then
    If IsDate(mskInicialPedidos) = False Or IsDate(mskFinalPedidos) = False Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (pedidos.data_compra >= CONVERT(DATETIME, '" & Format(mskInicialPedidos.Text, ocDATA) & "', 103)) AND (pedidos.data_compra <= CONVERT(DATETIME, '" & Format(mskFinalPedidos.Text, ocDATA) & "', 103)) ORDER BY pedidos.cod_pedido;"
ElseIf cboIndicePedidos.Text = "MENSAL" Then
    If cboMesPedidos.Text = "" Or cboAnoPedidos.Text = "" Then Exit Sub
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (MONTH(pedidos.data_compra) = " & cboMesPedidos.ListIndex + 1 & ") AND (YEAR(pedidos.data_compra) = " & cboAnoPedidos & ") ORDER BY pedidos.cod_pedido;"
End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGridPedidos r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'MOSTRAR A QUANTIDADE REGISTROS
   lblQuantPedidos.Caption = Format(totalRegistros, "00")
End Sub

Private Sub FormatarGridPedidos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridPedidos
       .Clear
       .Cols = 7
       .Rows = 2
           
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
            .TextMatrix(.Rows - 1, 1) = rTabela("var_codped")
            .TextMatrix(.Rows - 1, 2) = rTabela("status")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 4) = rTabela("var_Nome")
            .TextMatrix(.Rows - 1, 5) = rTabela("var_tipoPGTO")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("var_total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
'   lblValor.Caption = Format(SomaGrid(GridPedidos, 5), ocMONEY)
End Sub

Private Sub Command3_Click()

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

Private Sub mskSaida_KeyPress(KeyAscii As Integer)
mskSaida.Mask = "##/##/####"
End Sub

Private Sub TxtCodCliente_Change()
Dim TbClientes As New ADODB.Recordset

If TxtCodCliente.Text = "" Then Exit Sub

If cboTipoDest.Text = "FORNECEDOR" Then
    RsOpen TbClientes, "SELECT * FROM fornecedor WHERE codigo = " & Val(TxtCodCliente.Text)
Else
    RsOpen TbClientes, "SELECT * FROM cliente WHERE codigo = " & Val(TxtCodCliente.Text)
End If

If TbClientes.EOF And TbClientes.BOF Then
    'MsgBox "Código do cliente não foi localizado no sistema. Verifique.", vbCritical, "Online Commerce": Exit Sub
Else
    txtCliEndereco.Text = ValidateNull(TbClientes("endereco"))
    txtCliNum.Text = ValidateNull(TbClientes("numero"))
    txtCliBairro.Text = ValidateNull(TbClientes("bairro"))
    txtCliCidade.Text = ValidateNull(TbClientes("cidade"))
    txtCliUF.Text = ValidateNull(TbClientes("estado"))
    txtCliIBGE.Text = ValidateNull(TbClientes("CodigoIBGE"))
    
    If cboTipoDest.Text = "FORNECEDOR" Then
        txtCliCPF.Text = ValidateNull(TbClientes("cnpj"))
    Else
        txtCliCPF.Text = ValidateNull(TbClientes("cpf"))
    End If
    
    txtCliIE.Text = ValidateNull(TbClientes("ie"))
End If
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   'Exit Sub
   If KeyCode = 38 Then
      If GridNotasItens.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
      GridNotasItens.Row = iRow - 1
      GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      GridNotasItens_Click
   
   ElseIf KeyCode = 40 Then
      If GridNotasItens.Rows = GridNotasItens.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
      GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      GridNotasItens.Row = iRow + 1
      GridNotasItens_Click
   End If
End Sub
Private Sub txtEdit_LostFocus()
   GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
   txtEdit.Visible = False

AtualizarGrid_Itens
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
SSTab1.Tab = 0
SSTab2.Tab = 0
SSTab3.Tab = 0

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

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdTransmitir.Enabled = False
cmdImprimir.Enabled = False
cmdCancelarNota.Enabled = False
cmdConsultar.Enabled = False
frmNota.Enabled = False
'frmTransmissao.Enabled = False
frmItens.Enabled = False
lblTipoConsulta.Caption = "0"
End Sub

Private Sub GridNotas_DblClick()
'Clear_Controls
'LimparCamposItens
If cmdSalvar.Enabled = True Then
    MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub
Else
    RsOpen TbNotas, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal WHERE CodigoNota = " & GridNotas.TextMatrix(GridNotas.Row, 1)
    Load_Controls
    Frm_NF.Tab = 0
End If
End Sub

Private Sub GridNotasItens_Click()
Dim i As Integer

For i = 3 To 5
   If GridNotasItens.ColSel = i Then
      txtEdit.Move GridNotasItens.Left + GridNotasItens.CellLeft, GridNotasItens.Top + GridNotasItens.CellTop, GridNotasItens.CellWidth, GridNotasItens.CellHeight
      txtEdit.Text = GridNotasItens.TextMatrix(GridNotasItens.Row, GridNotasItens.Col)
      txtEdit.Visible = True
      txtEdit.SetFocus
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(txtEdit.Text)
      iRow = GridNotasItens.Row
      iCol = GridNotasItens.Col
   End If
Next
End Sub

Private Sub GridPedidos_DblClick()
If GridPedidos.TextMatrix(GridPedidos.Row, 2) = "SIM" Then MsgBox "Esse pedido já foi transformado em NFe!", vbInformation, "Online Commerce": Exit Sub
If ShowMsg("Deseja realmente transformar o pedido: " & GridPedidos.TextMatrix(GridPedidos.Row, 1) & " em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

txtCodPedido.Text = (GridPedidos.TextMatrix(GridPedidos.Row, 1))
    
'TransformarPedidoemNFE
GravarPedido

cmdNovo.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdImprimir.Enabled = False
cmdTransmitir.Enabled = True
'cmdCancelarNota.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True

Frm_NF.Tab = 0
End Sub


Private Sub Label26_Click()
'chkDesc.Value = 1
'chkDesc_Click
End Sub

Private Sub lblCodFabrica_Click()
chkCodBarra.Value = 1
chkCodBarra_Click
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.Rows - 1
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

Private Sub txtInfAdicionais_GotFocus()
txtInfAdicionais.SelStart = 0
txtInfAdicionais.SelLength = Len(txtInfAdicionais)
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

Private Sub Text11_GotFocus()
Text11.SelStart = 0
Text11.SelLength = Len(Text11)
End Sub


Private Sub txtVolMarca_GotFocus()
txtVolMarca.SelStart = 0
txtVolMarca.SelLength = Len(txtVolMarca)
End Sub


Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13)
End Sub


Private Sub txtCodObservacao_GotFocus()
txtCodObservacao.SelStart = 0
txtCodObservacao.SelLength = Len(txtCodObservacao)
End Sub

Private Sub txtValorOutrasDespesas_GotFocus()
txtValorOutrasDespesas.SelStart = 0
txtValorOutrasDespesas.SelLength = Len(txtValorOutrasDespesas)
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
txtValorSeguro.SelStart = 0
txtValorSeguro.SelLength = Len(txtValorSeguro)
End Sub

Private Sub Mostrar_Pedido(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then

Dim totalRegistros As Long

    'buscar Numero e codigo da nota (autopreenchimento)
    RsOpen TbNotaPedido, "SELECT *, " & _
                "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                "FROM NotaFiscal"
                
    If TbNotaPedido.RecordCount > 0 Then totalRegistros = TbNotaPedido.RecordCount
        
    'Clear_Controls
    
    'INICIO DO PREENCHIMENTOS DOS OBJETOS
    If TbNotaPedido.EOF And TbNotaPedido.BOF Then
        txtNumNota.Text = "1"
        txtCodNota.Text = "1"
    Else
        TbNotaPedido.MoveLast
        txtNumNota.Text = TbNotaPedido("NumeroNota") + 1
        txtCodNota.Text = TbNotaPedido("CodigoNota") + 1
    End If

    
'    txtNumNota = Format(rTabela("NumeroNota"), "@")
'    txtCodNota = Format(rTabela("CodigoNota"), "@")
    TxtCodCliente = Format(rTabela("COD_CLIENTE"), "@")
    CboCliente = Format(rTabela("varnome"), "@")
    'cboDestOperacao = Format(5102, "@")
    cboDestOperacao.Text = "1 - Operação Interna"
    'txtInfAdicionais = Format(rTabela("InformacoesAdicionais"), "@")
    cboNatureza = "VENDA DE MERCADORIA"
    'mskEmissao = Format(Date, "dd/mm/yyyy")
    'mskSaida = Format(Date, "dd/mm/yyyy")
    'mskHora = Format(Time(), "HH:MM:ss")
    txtCodTransporte = Format(0, "@")
    cboTransporte = Format(0, "@")
    txtPlaca = Format(0, "@")
    txtVolQuant = Format(0, "@")
    Text11 = Format(0, "@")
    txtVolMarca = Format(0, "@")
    Text13 = Format(0, "@")
    txtCodObservacao = Format(0, "@")
    
    txtValorSeguro = Format(0, "##,##0.00")
    txtValorOutrasDespesas = Format(0, "##,##0.00")
    txtValorFrete = Format(0, "##,##0.00")
    txtBaseICMS = Format(0, "##,##0.00")
    txtBaseICMSST = Format(0, "##,##0.00")
    txtVolPesoBruto = Format(0, "@")
    txtVolPesoLiquido = Format(0, "@")
    txtPlacaUF = Format(0, "@")
    cboModFrete = "9 - SEM FRETE"
    
    Text30 = Format(0, "@")
    Text31 = Format(0, "@")
    Text32 = Format(0, "@")
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
End If
End Sub

Private Sub txtDesc_Change()
'On Error GoTo Erro

If txtDesc.Text = "" Or txtValor.Text = "" Then
   txtDesc.Text = "0"
   SelectControl txtDesc
   Exit Sub
End If

Calcular_Desconto
Exit Sub
   
'Erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   txtDesc.Text = 0
End Sub

Private Sub Calcular_Desconto()
Dim varTotalUnid As Currency
Dim varTotalSemDesc As Currency
Dim varTotalComdesc As Currency
Dim varQuant As Double
Dim varDesc As Currency

If txtDesc.Text = "" Then txtDesc.Text = "0"
If txtQuant.Text = "" Then txtQuant.Text = "1": SelectControl txtDesc
If txtValor.Text = "" Then txtValor.Text = "0"

varTotalUnid = txtValor.Text

varDesc = txtDesc.Text
varQuant = txtQuant.Text
varTotalSemDesc = CCur(varTotalUnid) * CDbl(varQuant)
varTotalComdesc = CCur(varTotalUnid) * CDbl(varQuant) - CCur(varDesc)

If txtValor.Text = "0" Or txtValor.Text = "0,00" Then Exit Sub
If txtDesc.Text = "" Then txtDesc.Text = Format(0, ocMONEY)

If txtDesc.Text <> "0,00" Then
    txtSubTotal.Text = Format(varTotalComdesc, ocMONEY)
Else
    txtSubTotal.Text = Format(varTotalSemDesc, ocMONEY)
End If
End Sub

Private Sub txtDesc_GotFocus()
SelectControl txtDesc
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtDesc_LostFocus()
On Error GoTo erro

If txtDesc.Text = "" Or txtValor.Text = "" Then
   txtDesc.Text = 0
   SelectControl txtDesc
   Exit Sub
End If

Calcular_Desconto
txtDesc.Text = Format(txtDesc.Text, ocMONEY)
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
txtBaseICMSST.SelStart = 0
txtBaseICMSST.SelLength = Len(txtBaseICMSST)
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
Calcular_Desconto
End Sub

Private Sub txtSubTotal_GotFocus()
SelectControl txtSubTotal
End Sub


Private Sub txtValor_GotFocus()
SelectControl txtValor
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub

Private Sub txtValorFrete_GotFocus()
txtValorFrete.SelStart = 0
txtValorFrete.SelLength = Len(txtValorFrete)
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
End Sub

Private Sub txtCodBarra_GotFocus()
'txtCodBarra.SelStart = 0
'txtCodBarra.SelLength = Len(txtCodBarra)
SelectControl txtCodBarra
End Sub


Private Sub txtCodBarra_LostFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If lblTipoConsulta.Caption = "0" Or lblTipoConsulta.Caption = "1" Then
    If txtCodBarra.Text = "" Then
        lblTipoConsulta.Caption = "0"
        txtCodProduto.Text = ""
        cboDescricao.Locked = False
        cboDescricao.Text = ""
        txtUnid.Text = ""
        txtCFOP.Text = ""
        txtCST.Text = ""
        txtNCM.Text = ""
        txtICMS.Text = ""
        txtValor.Text = "0"
        txtSubTotal.Text = "0"
        Exit Sub
    End If



        sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante, quant_estoque, unid_medida, CFOP, NCM, ICMSCST, ICMSAliq  FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "') AND (ativo = 1);"
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r.BOF Then
           txtCodProduto.Text = r("var_codprod")
           
           If tipoEmpresa = 4 Then
               cboDescricao.Text = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante")) & " /  " & r("REF")
               'cboDescricao2.Text = ValidateNull(r("var_desc"))
           Else
              cboDescricao.Text = ValidateNull(r("var_desc"))
              txtUnid.Text = ValidateNull(r("unid_medida"))
              txtCFOP.Text = ValidateNull(r("CFOP"))
              txtCST.Text = ValidateNull(r("ICMSCST"))
              txtNCM.Text = ValidateNull(r("NCM"))
              txtICMS.Text = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
           End If
           
            lblTipoConsulta.Caption = "1"
            cboDescricao.Locked = True
        Else
           ShowMsg "Produto Inexistente!", vbCritical
           lblTipoConsulta.Caption = "0"
           txtCodBarra.Text = ""
           txtCodBarra.SetFocus
           Exit Sub
        End If
        
        MostrarValorVenda
        txtQuant.SetFocus
    'End If
End If
On Local Error Resume Next
End Sub

Private Sub lblTipoConsulta_Change()
If lblTipoConsulta.Caption = "1" Then
    txtCodBarra.BackColor = &HC0FFFF
    cboDescricao.BackColor = &HFFFFFF
ElseIf lblTipoConsulta.Caption = "2" Then
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HC0FFFF
Else
    txtCodBarra.BackColor = &HFFFFFF
    cboDescricao.BackColor = &HFFFFFF
    txtCodBarra.Locked = False
    cboDescricao.Locked = False
End If
End Sub
Private Sub txtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbClientes As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If TxtCodCliente.Text = "" Then
        TxtCodCliente.SetFocus
        Exit Sub
    End If
    RsOpen TbClientes, "SELECT * FROM cliente WHERE codigo = " & Val(TxtCodCliente.Text)
    If TbClientes.EOF And TbClientes.BOF Then
        MsgBox "Código do cliente não foi localizado no sistema. Verifique.", vbCritical, "Online Commerce": Exit Sub
    Else
        CboCliente.Text = TbClientes("nome")
        txtCliEndereco.Text = TbClientes("nome")
        txtCliNum.Text = TbClientes("nome")
        txtCliBairro.Text = TbClientes("nome")
        txtCliCidade.Text = TbClientes("nome")
        txtCliUF.Text = TbClientes("nome")
        txtCliCPF.Text = TbClientes("nome")
        txtCliIE.Text = TbClientes("nome")
        
        CboCliente.SetFocus
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Sub cboCliente_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub txtCodTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbTransportadora As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If txtCodTransporte.Text = "" Then
        'Frm_LTransportadora.Show 1
        Exit Sub
    End If
    RsOpen TbTransportadora, "select * from transportadora where codigo=" & Val(txtCodTransporte.Text)
    If TbTransportadora.BOF And TbTransportadora.BOF Then
        MsgBox "Não foi possivel localizar o código da transportadora. Favor verifique.", vbCritical, "Online Commerce": Exit Sub
    Else
        txtCodTransporte.Text = TbTransportadora("codigo")
        cboTransporte.Text = TbTransportadora("razao")
        txtCodTransporte.SetFocus
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
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
 If TbNotas("Denegada") Then
    MostraStatus.ForeColor = vbRed
    MostraStatus_F9$ = "Denegada"  'Deve retornar uma expressão caractere
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdImprimir.Enabled = True
    cmdConsultar.Enabled = True
 ElseIf TbNotas("Enviada") And Not TbNotas("Cancelada") Then
    MostraStatus.ForeColor = vbBlue
    MostraStatus_F9$ = "Transmitida/Autorizada"  'Deve retornar uma expressão caractere
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = True
    cmdImprimir.Enabled = True
    cmdConsultar.Enabled = True
 ElseIf TbNotas("Cancelada") Then
    MostraStatus.ForeColor = vbRed
    MostraStatus_F9$ = "Cancelada"  'Deve retornar uma expressão caractere
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdTransmitir.Enabled = False
    cmdCancelarNota.Enabled = False
    cmdImprimir.Enabled = True
    cmdConsultar.Enabled = True
 Else
    MostraStatus.ForeColor = vbBlack
    MostraStatus_F9$ = "Em Digitação"  'Deve retornar uma expressão caractere
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdTransmitir.Enabled = True
    cmdCancelarNota.Enabled = False
    cmdImprimir.Enabled = False
    cmdConsultar.Enabled = False
 End If
End Function

Private Sub LimparGridNotas()
   Dim i As Integer
   
   With GridNotas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 500
      .ColWidth(2) = 1000
      .ColWidth(3) = 1100
      .ColWidth(4) = 2000
      .ColWidth(5) = 3000
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
      .Rows = .Rows + 1
      
      i = i + 1
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
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
   Dim i As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 2000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "CST"
      .TextMatrix(0, 4) = "UND"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "SUBTOTAL"
      
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
      .ColAlignment(5) = 2
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      .Rows = .Rows + 1
      
      i = i + 1
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
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
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 500
      .ColWidth(2) = 1000
      .ColWidth(3) = 1100
      .ColWidth(4) = 2000
      .ColWidth(5) = 3000
      .ColWidth(6) = 1500
      .ColWidth(7) = 2000
      
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
      .ColAlignment(7) = 2
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("CodigoNota")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("NumeroNota"), "000000")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("DataEmissao"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 4) = rTabela("NaturezaOperacao")
            .TextMatrix(.Rows - 1, 5) = rTabela("RazaoSocial")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("ValorNota"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = rTabela("Status")
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
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
If txtQuant.Text = "" Or txtValor.Text = "" Then
   txtQuant.Text = "1"
   SelectControl txtQuant
   Exit Sub
End If

Calcular_Desconto
Exit Sub
End Sub

Private Sub txtQuant_GotFocus()
SelectControl txtQuant
End Sub


Private Sub txtQuant_LostFocus()
'Calcular_Total
Calcular_Desconto
End Sub


Private Sub txtValor_Change()
'Calcular_Total
Calcular_Desconto
End Sub


Private Sub txtValor_LostFocus()
'Calcular_Total
Calcular_Desconto
End Sub


