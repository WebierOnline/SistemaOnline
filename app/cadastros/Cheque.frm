VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Cheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHEQUE"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
   Icon            =   "Cheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4755
      Left            =   60
      TabIndex        =   18
      Top             =   1080
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   8387
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   450
      TabMaxWidth     =   2646
      WordWrap        =   0   'False
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "Cheque.frx":23D2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmCadastro"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "Cheque.frx":23EE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTotal"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Grid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   1035
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   10995
         Begin VB.ComboBox cboMES 
            Height          =   315
            ItemData        =   "Cheque.frx":240A
            Left            =   4320
            List            =   "Cheque.frx":240C
            TabIndex        =   44
            Top             =   420
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   6120
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   420
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboConsProprietario 
            Height          =   315
            ItemData        =   "Cheque.frx":240E
            Left            =   4320
            List            =   "Cheque.frx":2410
            TabIndex        =   39
            Top             =   420
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.ComboBox cboOrganizar 
            Height          =   315
            ItemData        =   "Cheque.frx":2412
            Left            =   2160
            List            =   "Cheque.frx":2414
            TabIndex        =   37
            Top             =   420
            Width           =   2115
         End
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            ItemData        =   "Cheque.frx":2416
            Left            =   60
            List            =   "Cheque.frx":2418
            TabIndex        =   35
            Top             =   420
            Width           =   2055
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir 
            Height          =   555
            Left            =   8160
            TabIndex        =   41
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "Cheque.frx":241A
            PICN            =   "Cheque.frx":2436
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
            Height          =   555
            Left            =   9540
            TabIndex        =   42
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   979
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
            MICON           =   "Cheque.frx":41C8
            PICN            =   "Cheque.frx":41E4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdConsData 
            Height          =   315
            Left            =   5460
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   420
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
            MICON           =   "Cheque.frx":5F76
            PICN            =   "Cheque.frx":5F92
            PICH            =   "Cheque.frx":82E5
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
            Left            =   4320
            TabIndex        =   47
            Top             =   420
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label lblConsData 
            BackStyle       =   0  'Transparent
            Caption         =   "Pré-Datado"
            Height          =   255
            Left            =   4320
            TabIndex        =   48
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblCONmes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&scolha o mês/ano:"
            Height          =   195
            Left            =   4320
            TabIndex        =   45
            Top             =   180
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label lblConsProprietario 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   4320
            TabIndex        =   40
            Top             =   180
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   2160
            TabIndex        =   38
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Critério:"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   180
            Width           =   525
         End
      End
      Begin VB.Frame frmCadastro 
         Height          =   4155
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   10995
         Begin ChamaleonBtn.chameleonButton cmdCal1 
            Height          =   315
            Left            =   7740
            TabIndex        =   32
            TabStop         =   0   'False
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
            MICON           =   "Cheque.frx":A638
            PICN            =   "Cheque.frx":A654
            PICH            =   "Cheque.frx":C9A7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboOperacao 
            Height          =   315
            ItemData        =   "Cheque.frx":ECFA
            Left            =   120
            List            =   "Cheque.frx":ECFC
            TabIndex        =   1
            Top             =   420
            Width           =   2235
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            ItemData        =   "Cheque.frx":ECFE
            Left            =   120
            List            =   "Cheque.frx":ED00
            TabIndex        =   2
            Top             =   1140
            Width           =   2295
         End
         Begin VB.ComboBox cboProprietario 
            Height          =   315
            ItemData        =   "Cheque.frx":ED02
            Left            =   120
            List            =   "Cheque.frx":ED04
            TabIndex        =   9
            Top             =   1920
            Width           =   9015
         End
         Begin VB.TextBox txtCodOperacao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAgencia 
            Height          =   315
            Left            =   2460
            TabIndex        =   3
            Top             =   1140
            Width           =   1515
         End
         Begin VB.TextBox txtCC 
            Height          =   315
            Left            =   4020
            TabIndex        =   4
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtNumero 
            Height          =   315
            Left            =   5460
            TabIndex        =   5
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8040
            TabIndex        =   7
            Top             =   1140
            Width           =   1455
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   315
            Left            =   10560
            TabIndex        =   21
            TabStop         =   0   'False
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
            MICON           =   "Cheque.frx":ED06
            PICN            =   "Cheque.frx":ED22
            PICH            =   "Cheque.frx":11075
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskEmissao 
            Height          =   315
            Left            =   6720
            TabIndex        =   6
            Top             =   1140
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskPredatado 
            Height          =   315
            Left            =   9540
            TabIndex        =   8
            Top             =   1140
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Operação"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   5460
            TabIndex        =   30
            Top             =   900
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   900
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            Height          =   195
            Left            =   2460
            TabIndex        =   28
            Top             =   900
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   4020
            TabIndex        =   27
            Top             =   900
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Proprietario"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label21 
            Caption         =   "Emissão"
            Height          =   255
            Left            =   6720
            TabIndex        =   25
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Pré-Datado"
            Height          =   255
            Left            =   9540
            TabIndex        =   24
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   8040
            TabIndex        =   23
            Top             =   900
            Width           =   360
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3135
         Left            =   60
         TabIndex        =   19
         Top             =   1380
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   5530
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblTotal 
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
         Left            =   10920
         TabIndex        =   33
         Top             =   4500
         Width           =   225
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12660
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   13545
      TabIndex        =   15
      Top             =   60
      Width           =   13575
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   300
         Picture         =   "Cheque.frx":133C8
         Top             =   -320
         Width           =   1500
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CHEQUE"
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
         Left            =   1980
         TabIndex        =   16
         Top             =   240
         Width           =   1320
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   17
      Top             =   5925
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19870
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:19"
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
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   615
      Left            =   11400
      TabIndex        =   11
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Cheque.frx":1452E
      PICN            =   "Cheque.frx":1454A
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
      Left            =   11400
      TabIndex        =   12
      Top             =   3300
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Cheque.frx":162DC
      PICN            =   "Cheque.frx":162F8
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
      Left            =   11400
      TabIndex        =   13
      Top             =   3960
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Cheque.frx":1808A
      PICN            =   "Cheque.frx":180A6
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
      Left            =   11400
      TabIndex        =   10
      Top             =   1980
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Cheque.frx":19E38
      PICN            =   "Cheque.frx":19E54
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
      Left            =   11400
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Cheque.frx":1BBE6
      PICN            =   "Cheque.frx":1BC02
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
Attribute VB_Name = "Cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private printSQL As String
Private Function Inserir_Dados() As Boolean
Dim sSQL As String

sSQL = "INSERT INTO Cheque (codigo, cod_operacao, operacao, banco, agencia, cc, numero, proprietario, valor, emissao, predatado, status) VALUES (" & _
   txtCodigo.Text & ", " & txtCodOperacao.Text & ", '" & cboOperacao.Text & "', '" & cboBanco.Text & "', '" & txtAgencia.Text & "', '" & txtCC.Text & "', " & txtNumero.Text & ", '" & cboProprietario.Text & "', " & Replace(CCur(txtValor.Text), ",", ".") & ", CONVERT(DATETIME, '" & Format$(mskEmissao.Text, ocDATA) & "', 103), CONVERT(DATETIME, '" & Format$(mskPredatado.Text, ocDATA) & "', 103), 0);"

Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

sSQL = "UPDATE Cheque SET cod_operacao= " & txtCodOperacao.Text & ", operacao = '" & cboOperacao.Text & "', banco = '" & cboBanco.Text & "', agencia = '" & txtAgencia.Text & "', cc = '" & txtCC.Text & "', Numero = " & txtNumero.Text & ", proprietario = '" & cboProprietario.Text & "', Valor = " & Replace(CCur(txtValor.Text), ",", ".") & ", emissao = CONVERT(DATETIME, '" & Format$(mskEmissao.Text, ocDATA) & "', 103), predatado = CONVERT(DATETIME, '" & Format$(mskPredatado.Text, ocDATA) & "', 103) WHERE (codigo = " & txtCodigo.Text & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS codigo FROM Cheque;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodigo.Text = r("codigo") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer

With Grid
   .Clear
   .Cols = 12
   .Rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 950
   .ColWidth(3) = 800
   .ColWidth(4) = 2000
   .ColWidth(5) = 1700
   .ColWidth(6) = 800
   .ColWidth(7) = 800
   .ColWidth(8) = 900
   .ColWidth(9) = 950
   .ColWidth(10) = 1000
   .ColWidth(11) = 900
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "Operação"
   .TextMatrix(0, 3) = "Emissão"
   .TextMatrix(0, 4) = "Proprietário"
   .TextMatrix(0, 5) = "Banco"
   .TextMatrix(0, 6) = "Núm."
   .TextMatrix(0, 7) = "Agência"
   .TextMatrix(0, 8) = "Conta"
   .TextMatrix(0, 9) = "Valor"
   .TextMatrix(0, 10) = "Datado"
   .TextMatrix(0, 11) = "Situação"
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.Rows - 1, 2) = rTabela("operacao")
         .TextMatrix(.Rows - 1, 3) = Format(rTabela("emissao"), "dd/mm/yy")
         .TextMatrix(.Rows - 1, 4) = rTabela("proprietario")
         .TextMatrix(.Rows - 1, 5) = rTabela("banco")
         .TextMatrix(.Rows - 1, 6) = rTabela("numero")
         .TextMatrix(.Rows - 1, 7) = rTabela("agencia")
         .TextMatrix(.Rows - 1, 8) = rTabela("cc")
         .TextMatrix(.Rows - 1, 9) = Format(rTabela("valor"), ocMONEY)
         .TextMatrix(.Rows - 1, 10) = Format(rTabela("predatado"), "dd/mm/yy")
         .TextMatrix(.Rows - 1, 11) = rTabela("var_status")
         rTabela.MoveNext
         
         .Rows = .Rows + 1
         i = i + 1
      Loop
   End If
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .Rows - 1
      .Row = i
      .Col = 9
      '.CellForeColor = &HC0&
      .CellFontBold = True
   Next

   For i = 1 To .Rows - 1
      .Row = i
      .Col = 10
      .CellForeColor = &HC0&
      .CellFontBold = True
   Next
   
   .Rows = .Rows - 1
   .Redraw = True

lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End With
End Sub

Private Sub Limpar_Objetos()
txtCodigo.Text = ""
cboProprietario.Text = ""
txtCodOperacao.Text = ""
cboOperacao.Text = ""
txtNumero.Text = ""
cboBanco.Text = ""
txtAgencia.Text = ""
txtCC.Text = ""
txtValor.Text = ""
cboProprietario.Text = ""
mskEmissao.Mask = ""
mskEmissao.Text = ""
mskPredatado.Mask = ""
mskPredatado.Text = ""
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


Private Sub cboBanco_GotFocus()
cboBanco.Clear
cboBanco.AddItem "BANCO DO BRASIL"
cboBanco.AddItem "BRADESCO"
cboBanco.AddItem "BANCO DO NORDESTE"
cboBanco.AddItem "SANTANDER"
cboBanco.AddItem "ITAU"
moCombo.AttachTo cboBanco
End Sub


Private Sub cboConsProprietario_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
cboConsProprietario.Clear

sSQL = "SELECT DISTINCT proprietario FROM cheque ORDER BY proprietario;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboConsProprietario.AddItem r("proprietario")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboConsProprietario
End Sub


Private Sub cboCriterio_GotFocus()
cboCriterio.Clear
cboCriterio.AddItem "TODOS"
cboCriterio.AddItem "PROPRIETÁRIO"
cboCriterio.AddItem "MENSAL"
cboCriterio.AddItem "DATA"
moCombo.AttachTo cboCriterio
End Sub


Private Sub cboCriterio_LostFocus()
If cboCriterio.Text = "PROPRIETÁRIO" Then
    lblConsProprietario.Visible = True
    cboConsProprietario.Visible = True
    lblConsData.Visible = False
    mskConsData.Visible = False
    cmdConsData.Visible = False
    lblCONmes.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
ElseIf cboCriterio.Text = "MENSAL" Then
    lblConsProprietario.Visible = False
    cboConsProprietario.Visible = False
    lblConsData.Visible = False
    mskConsData.Visible = False
    cmdConsData.Visible = False
    lblCONmes.Visible = True
    cboMes.Visible = True
    cboAno.Visible = True
ElseIf cboCriterio.Text = "DATA" Then
    lblConsProprietario.Visible = False
    cboConsProprietario.Visible = False
    lblConsData.Visible = True
    mskConsData.Visible = True
    cmdConsData.Visible = True
    lblCONmes.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
Else
    lblConsProprietario.Visible = False
    cboConsProprietario.Visible = False
    lblConsData.Visible = False
    mskConsData.Visible = False
    cmdConsData.Visible = False
    lblCONmes.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
End If
End Sub


Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear

For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
End Sub


Private Sub cboOperacao_GotFocus()
cboOperacao.Clear
cboOperacao.AddItem "AVULSO"
cboOperacao.AddItem "SUPRIMENTO"
cboOperacao.AddItem "PEDIDO"
cboOperacao.AddItem "O.S."
moCombo.AttachTo cboOperacao
End Sub


Private Sub cboOperacao_LostFocus()
If cboOperacao.Text = "AVULSO" Then txtCodOperacao.Text = "0"
End Sub


Private Sub cboOrganizar_GotFocus()
cboOrganizar.Clear
cboOrganizar.AddItem "PRÉ-DATADO"
cboOrganizar.AddItem "PROPRIETÁRIO"
cboOrganizar.AddItem "EMISSÃO"
moCombo.AttachTo cboOrganizar
End Sub


Private Sub cboProprietario_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String

itemAtual = cboProprietario.Text
cboProprietario.Clear

sSQL = "SELECT DISTINCT proprietario FROM cheque ORDER BY proprietario;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboProprietario.AddItem r("proprietario")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboProprietario.Text = itemAtual
moCombo.AttachTo cboProprietario
End Sub


Private Sub cboProprietario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub chameleonButton1_Click()
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

mskPredatado = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Or cboProprietario.Text = "" Then Exit Sub

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
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

mskEmissao = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos
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
Dim sSQL As String
Dim bRet As Boolean

If txtCodigo.Text = "" Then Exit Sub

If ShowMsg("Excluir esse Cheque?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

sSQL = "DELETE FROM Cheque WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdExibir_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If cboCriterio.Text = "" Then Exit Sub

Dim INDICE As String
If cboOrganizar.Text = "PRÉ-DATADO" Then
   INDICE = "predatado "
ElseIf cboOrganizar.Text = "PROPRIETÁRIO" Then
   INDICE = "proprietario "
ElseIf cboOrganizar.Text = "EMISSÃO" Then
   INDICE = "emissao "
Else
   INDICE = "emissao "
End If

If cboCriterio.Text = "TODOS" Then
    sSQL = "SELECT *, CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_Status FROM Cheque ORDER BY  " & INDICE
ElseIf cboCriterio.Text = "PROPRIETÁRIO" Then
    sSQL = "SELECT *, CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_Status FROM Cheque where (proprietario = '" & cboConsProprietario.Text & "') ORDER BY  " & INDICE
ElseIf cboCriterio.Text = "MENSAL" Then
    If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
    sSQL = "SELECT *, CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_Status FROM Cheque where  (MONTH(predatado) = " & cboMes.ListIndex + 1 & ") AND (YEAR(predatado) = " & cboAno & ") ORDER BY  " & INDICE
ElseIf cboCriterio.Text = "DATA" Then
    If mskConsData.Text = "" Then Exit Sub
    sSQL = "SELECT *, CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_Status FROM Cheque where (predatado = CONVERT(DATETIME, '" & Format(mskConsData, ocDATA) & "', 103))ORDER BY  " & INDICE
End If

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub
Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.Rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub cmdImprimir_Click()
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini
Dim r As ADODB.Recordset

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Cheque.Relatorio.Recordset = r
'REL_Cheque.dfQuant.Caption = "QUANTIDADE: " & txtCONquant.Text
REL_Cheque.dfTotal.Caption = "TOTAL: " & lblTotal.Caption
REL_Cheque.lblTitulo.Caption = "RELATÓRIO - CONTAS À RECEBER/RECEBIDO"

'If cboFiltro.Text = "TODOS" Then
'   REL_Cheque.dfTipo.Caption = "Tipo: Todos os registros"
'ElseIf cboFiltro.Text = "PERIODO" Then
'   REL_Cheque.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
'ElseIf cboFiltro.Text = "MÊS" Then
'   REL_Cheque.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
'ElseIf cboFiltro.Text = "CLIENTE" Then
'   REL_Cheque.dfTipo.Caption = "Cliente = " & cboNome.Text
'Else
'   REL_Cheque.dfTipo.Caption = "Tipo:"
'End If

REL_Cheque.Relatorio.NomeImpressora = var_Impressora
REL_Cheque.Relatorio.Ativar
Unload REL_Cheque

Me.Show 1

End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
Auto_Numeracao
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
frmCadastro.Enabled = True
cboOperacao.SetFocus
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo TrataErro

If txtCodigo.Text = "" Or cboProprietario.Text = "" Then Exit Sub

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
MostrarGrid
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdNovo.Enabled = True
frmCadastro.Enabled = False
SSTab1.Tab = 0
End Sub

Private Sub Mostrar_Cheque(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
   txtCodOperacao.Text = rTabela("cod_operacao")
   cboOperacao.Text = rTabela("operacao")
   cboProprietario.Text = rTabela("proprietario")
   cboBanco.Text = rTabela("banco")
   txtAgencia.Text = rTabela("agencia")
   txtCC.Text = rTabela("cc")
   txtNumero.Text = rTabela("numero")
   txtValor.Text = Format(rTabela("valor"), ocMONEY)
   mskEmissao.Text = Format(rTabela("emissao"), "dd/mm/yy")
   mskPredatado.Text = Format(rTabela("predatado"), "dd/mm/yy")
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub MostrarGrid()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT *, CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS var_Status FROM Cheque ORDER BY predatado;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub mskConsData_GotFocus()
If mskConsData.Text = "" Then mskConsData.Text = Format(Date, "dd/mm/yy")
SelectControl mskConsData
End Sub


Private Sub mskConsData_KeyPress(KeyAscii As Integer)
mskConsData.Mask = "##/##/##"
End Sub


Private Sub mskEmissao_GotFocus()
If mskEmissao.Text = "" Then mskEmissao.Text = Format(Date, "dd/mm/yy")
SelectControl mskEmissao
End Sub


Private Sub mskEmissao_KeyPress(KeyAscii As Integer)
mskEmissao.Mask = "##/##/##"
End Sub


Private Sub mskPredatado_GotFocus()
If mskPredatado.Text = "" Then mskPredatado.Text = Format(Date, "dd/mm/yy")
SelectControl mskPredatado
End Sub


Private Sub mskPredatado_KeyPress(KeyAscii As Integer)
mskPredatado.Mask = "##/##/##"
End Sub


Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

If cmdAlterar.Enabled = True Then
   sSQL = "SELECT * FROM cheque WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then Mostrar_Cheque r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If

SSTab1.Tab = 0
End Sub

Private Sub txtValor_GotFocus()
SelectControl txtValor
End Sub

Private Sub txtValor_LostFocus()
If txtValor.Text = "" Then
   txtValor.Text = Format(0, ocMONEY)
Else
   txtValor.Text = Format(txtValor, ocMONEY)
End If
End Sub


