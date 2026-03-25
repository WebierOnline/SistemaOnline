VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Caixa_Controle_semOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIXA"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "Caixa_Controle.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSetor 
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
      Height          =   975
      Left            =   4260
      TabIndex        =   47
      Top             =   7380
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ComboBox cboSetor 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   540
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdOKSetor 
         Height          =   315
         Left            =   1980
         TabIndex        =   49
         Top             =   540
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Setor"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame frmSenha 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4560
      TabIndex        =   44
      Top             =   6540
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin ChamaleonBtn.chameleonButton cmdSenha 
         Height          =   315
         Left            =   1500
         TabIndex        =   46
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":23EE
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
   Begin VB.CheckBox chkMostrarPrazo 
      Caption         =   "Mostrar À Prazo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   42
      Top             =   8580
      Width           =   1995
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Prazo 
      Height          =   2520
      Left            =   2940
      TabIndex        =   41
      Top             =   6180
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4445
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "SALDO:"
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
      Left            =   7920
      TabIndex        =   37
      Top             =   8940
      Width           =   3555
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   3315
      End
      Begin VB.CheckBox chkCartao 
         Caption         =   "C/ Cartão"
         Height          =   195
         Left            =   1200
         TabIndex        =   39
         Top             =   780
         Width           =   1095
      End
      Begin VB.CheckBox chkTroco 
         Caption         =   "C/ Troco"
         Height          =   195
         Left            =   2340
         TabIndex        =   38
         Top             =   780
         Width           =   1095
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Data"
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
      Left            =   120
      TabIndex        =   33
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   1200
         TabIndex        =   52
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
         MICON           =   "Caixa_Controle.frx":240A
         PICN            =   "Caixa_Controle.frx":2426
         PICH            =   "Caixa_Controle.frx":4779
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdOKData 
         Height          =   315
         Left            =   1560
         TabIndex        =   34
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":6ACC
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
         TabIndex        =   36
         Top             =   540
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         Caption         =   "Data do Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame frmMaquina 
      Caption         =   "Maquina"
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
      Left            =   0
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ComboBox cboMaquina 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   540
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdMaqOK 
         Height          =   315
         Left            =   1980
         TabIndex        =   31
         Top             =   540
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Caixa_Controle.frx":6AE8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Maquina"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.TextBox txtEntrada 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Top             =   8160
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame frmTroco 
      Caption         =   "Troco"
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
      Left            =   60
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   2835
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1395
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarTroco 
         Height          =   375
         Left            =   1620
         TabIndex        =   17
         Top             =   480
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
         MICON           =   "Caixa_Controle.frx":6B04
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhes do Caixa"
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
      Height          =   2835
      Left            =   7920
      TabIndex        =   7
      Top             =   6120
      Width           =   3555
      Begin VB.TextBox txtTotalAvulso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtTotalCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTotalCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtTotalTroco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTotalDinheiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "DEPOS. E TRANSF.:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CHEQUE:"
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
         Left            =   1200
         TabIndex        =   24
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "DINHEIRO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TROCO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CARTÃO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   10
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SAÍDA:"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   2100
         Width           =   630
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4995
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8811
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   11445
      TabIndex        =   4
      Top             =   0
      Width           =   11475
      Begin VB.Image Image1 
         Height          =   825
         Left            =   60
         Picture         =   "Caixa_Controle.frx":6B20
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CAIXA"
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
         Left            =   1260
         TabIndex        =   5
         Top             =   300
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controle"
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
      Left            =   60
      TabIndex        =   3
      Top             =   8940
      Width           =   7815
      Begin ChamaleonBtn.chameleonButton cmdMostrar 
         Height          =   675
         Left            =   60
         TabIndex        =   0
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1191
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
         MICON           =   "Caixa_Controle.frx":D261
         PICN            =   "Caixa_Controle.frx":D27D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirCaixa 
         Height          =   675
         Left            =   1980
         TabIndex        =   1
         Top             =   300
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Abrir Caixa"
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
         MICON           =   "Caixa_Controle.frx":DB57
         PICN            =   "Caixa_Controle.frx":DB73
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
         Height          =   675
         Left            =   5820
         TabIndex        =   2
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1191
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
         MICON           =   "Caixa_Controle.frx":DFDA
         PICN            =   "Caixa_Controle.frx":DFF6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdTroco 
         Height          =   675
         Left            =   3900
         TabIndex        =   13
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Troco"
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
         MICON           =   "Caixa_Controle.frx":E310
         PICN            =   "Caixa_Controle.frx":E32C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFecharCaixa 
         Height          =   675
         Left            =   1980
         TabIndex        =   20
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Fechar Caixa"
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
         MICON           =   "Caixa_Controle.frx":E45F
         PICN            =   "Caixa_Controle.frx":E47B
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
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1920
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirDetalhado 
      Default         =   -1  'True
      Height          =   675
      Left            =   360
      TabIndex        =   12
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1191
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
      MICON           =   "Caixa_Controle.frx":E8C8
      PICN            =   "Caixa_Controle.frx":E8E4
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
      TabIndex        =   51
      Top             =   10005
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10954
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Maquina"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Setor"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:12"
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Data"
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
      Height          =   255
      Left            =   2940
      TabIndex        =   43
      Top             =   8700
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblAviso1 
      AutoSize        =   -1  'True
      Caption         =   "Dê um duplo-clique na linha da venda para ver os produtos."
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
      Height          =   390
      Left            =   60
      TabIndex        =   28
      Top             =   6180
      Width           =   2835
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Caixa_Controle_semOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private printSQL As String
Private moCombo As cComboHelper
 
Private Sub Mostrar_Saldo()
   Dim var_Troco As Currency
   Dim var_Dinheiro As Currency
   Dim var_Cheque As Currency
   Dim var_Parcela As Currency
   Dim var_Cartao As Currency
   Dim var_Saida As Currency
   Dim var_Total As Currency
   Dim var_Entradas As Currency
   
   'Inicializa as variáveis
   var_Troco = 0
   var_Dinheiro = 0
   var_Cheque = 0
   var_Parcela = 0
   var_Cartao = 0
   var_Saida = 0
   
   If txtTotalTroco.Text <> "" Then var_Troco = txtTotalTroco.Text
   If txtTotalDinheiro.Text <> "" Then var_Dinheiro = txtTotalDinheiro.Text
   If txtTotalCartao.Text <> "" Then var_Cartao = txtTotalCartao.Text
   If txtTotalCheque.Text <> "" Then var_Cheque = txtTotalCheque.Text
   If txtTotalAvulso.Text <> "" Then var_Parcela = txtTotalAvulso.Text
   If txtSaida.Text <> "" Then var_Saida = txtSaida.Text
   
   If chkCartao.Value = Unchecked And chkTroco.Value = Unchecked Then
      var_Total = var_Dinheiro + var_Cheque + var_Parcela - var_Saida
      txtSaldo.Text = Format(var_Total, ocMONEY)
   
   ElseIf chkCartao.Value = Unchecked And chkTroco.Value = Checked Then
      var_Total = var_Troco + var_Dinheiro + var_Cheque + var_Parcela - var_Saida
      txtSaldo.Text = Format(var_Total, ocMONEY)
   
   ElseIf chkCartao.Value = Checked And chkTroco.Value = Unchecked Then
      var_Total = var_Cartao + var_Dinheiro + var_Cheque + var_Parcela - var_Saida
      txtSaldo.Text = Format(var_Total, ocMONEY)
   
   ElseIf chkCartao.Value = Checked And chkTroco.Value = Checked Then
      var_Total = var_Troco + var_Cartao + var_Dinheiro + var_Cheque + var_Parcela - var_Saida
      txtSaldo.Text = Format(var_Total, ocMONEY)
   End If
   
   var_Entradas = var_Cartao + var_Dinheiro + var_Cheque + var_Parcela
   txtEntrada.Text = Format(var_Entradas, ocMONEY)
End Sub

Private Sub SomaFlexCheque()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 5) = "CHEQUE" And IsNumeric(.TextMatrix(i, 6)) Then
            soma = soma + CCur(.TextMatrix(i, 6))
         End If
      Next
   End With
   
   txtTotalCheque.Text = Format(soma, ocMONEY)
   
errorhandeler:
End Sub

Private Sub SomaFlexParcela()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 5) = "DEPOSITO" Or .TextMatrix(i, 5) = "TRANSFERENCIA" And IsNumeric(.TextMatrix(i, 6)) Then
            soma = soma + CCur(.TextMatrix(i, 6))
         End If
      Next
   End With
   
   txtTotalAvulso.Text = Format(soma, ocMONEY)
   
errorhandeler:
End Sub



Private Sub SomaFlexSaida()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 3) = "SANGRIA" And IsNumeric(.TextMatrix(i, 7)) Then
            soma = soma + CCur(.TextMatrix(i, 7))
         End If
      Next
   End With
   
   txtSaida.Text = Format(soma, "#,##0.00")
   
errorhandeler:
End Sub

Private Sub SomaFlexCartao()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid
      For i = 1 To .Rows - 1
         If Left(.TextMatrix(i, 5), 6) = "CARTAO" And IsNumeric(.TextMatrix(i, 6)) Then
            soma = soma + CCur(.TextMatrix(i, 6))
         End If
      Next
   End With
   
   txtTotalCartao.Text = Format(soma, "#,##0.00")
   
errorhandeler:
End Sub

Private Sub SomaFlexDinheiro()
   On Error GoTo errorhandeler
   Dim soma As Currency
   Dim i As Integer
   
   soma = 0
   With Grid
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 5) = "DINHEIRO" And IsNumeric(.TextMatrix(i, 6)) Then
            soma = soma + CCur(.TextMatrix(i, 6))
         End If
      Next
   End With
   
   txtTotalDinheiro.Text = Format(soma, "#,##0.00")
   
errorhandeler:
End Sub

Private Sub Limpa_Tudo()
   'txtQuantEnt.Text = ""
   'txtQuantCyb.Text = ""
   'txtQuantSaida.Text = ""
   'txtTotMens.Text = ""
   'txtTotCyb.Text = ""
   'txtTotSaida.Text = ""
   txtTotalDinheiro.Text = ""
   txtSaida.Text = ""
   'txtTotal.Text = ""
   'txtQuantGraf.Text = ""
   'txtTotGraf.Text = ""
   'txtTotDiv.Text = ""
   'txtQuantSai.Text = ""
End Sub

Private Sub Mostrar_Troco()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim TROCO As Currency
   
   TROCO = 0
   sSQL = "SELECT * FROM caixa_troco WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then TROCO = r("valor")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   txtTotalTroco.Text = Format(TROCO, ocMONEY)
End Sub

Private Sub cboMaquina_GotFocus()
   cboMaquina.Clear
   cboMaquina.AddItem "CAIXA01"
   cboMaquina.AddItem "CAIXA02"
   cboMaquina.AddItem "CAIXA03"
   cboMaquina.AddItem "TODOS"
   moCombo.AttachTo cboMaquina
End Sub

Private Sub cboSetor_GotFocus()
   cboSetor.Clear
   cboSetor.AddItem "BALCAO"
   cboSetor.AddItem "OFICINA"
   cboSetor.AddItem "TODOS"
   moCombo.AttachTo cboSetor
End Sub

Private Sub chkCartao_Click()
   Mostrar_Saldo
End Sub

Private Sub chkMostrarPrazo_Click()
   If chkMostrarPrazo.Value = Checked Then
      Grid_Prazo.Visible = True
      lblTotalPrazo.Visible = True
      chkMostrarPrazo.Caption = "Ocultar À Prazo"
   Else
      Grid_Prazo.Visible = False
      lblTotalPrazo.Visible = False
      chkMostrarPrazo.Caption = "Mostrar À Prazo"
   End If
End Sub

Private Sub chkTroco_Click()
   Mostrar_Saldo
End Sub

Private Sub cmdAbrirCaixa_Click()
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   'cmdMostrar_Click
   
   If cmdAbrirCaixa.Visible = True Then
      frmSenha.Visible = True
      txtSenha.SetFocus
   End If
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

Private Sub cmdExibirDetalhado_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim Ent_Parcelas As Currency
   Dim Ent_Entradas As Currency
   Dim Soma_Entradas As Currency
   Dim Soma_Saidas As Currency
   Dim SALDO_FINAL As Currency
   
   If Not IsDate(mskData) Then Exit Sub
   
   'MOSTRAR As ENTRADAS
   sSQL = "SELECT parcelas.hora AS varHora, cliente.nome AS varCliente, parcelas.valor_final AS varValorLanc, " & _
      "CASE cod_os WHEN 0 THEN 'BALCÃO' ELSE 'OFICINA' END AS var_tipo, 0 AS varValorSaida, parcelas.valor_final AS campo04, " & _
      "parcelas.pagamento, pedidos.cod_pedido, pedidos.cod_cliente, cliente.codigo " & _
      "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
      "WHERE (parcelas.pagamento = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & _
      " UNION "
   
   sSQL = sSQL & "SELECT parcelas.hora AS varHora, cliente.nome AS varCliente, parcelas.valor_final AS varValorLanc, 0 AS var_tipo, " & _
      "0 AS varValorSaida, parcelas.valor_final AS campo04, parcelas.pagamento, os.cod_os, os.cod_cliente, cliente.codigo " & _
      "FROM cliente INNER JOIN os ON cliente.codigo = os.cod_cliente INNER JOIN parcelas ON parcelas.cod_os = os.cod_os " & _
      "WHERE (parcelas.pagamento = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & _
      " UNION "
   
   sSQL = sSQL & "SELECT hora AS varHora, descricao AS varCliente, valor AS varValorLanc, 0 AS var_tipo, 0 AS varValorSaida, " & _
      "valor AS campo04, cod_haver, setor, data, codigo FROM caixa_entrada " & _
      "WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & _
      " UNION "
   
   sSQL = sSQL & "SELECT caixa_saida.hora AS varHora, caixa_saida.descricao AS varCliente, 0 AS varValorLanc, setor as var_tipo, " & _
      "caixa_saida.valor AS varValorSaida, (0 - caixa_saida.valor) AS campo04, cod_haver, setor, data, codigo " & _
      "FROM caixa_saida WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) "     'saída do caixa
   
   sSQL = sSQL & "ORDER BY hora;"
   
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGridEntradaDetalhado r
   If r.State <> 0 Then r.Close
   Set r = Nothing

   'SOMAR ENTRADAS
   sSQL = "SELECT ISNULL(SUM(valor_final), 0) AS var_total FROM parcelas WHERE (pagamento = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103));"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Ent_Parcelas = r("var_total")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   sSQL = "SELECT ISNULL(SUM(valor), 0) AS var_total FROM caixa_entrada WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103));"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Ent_Entradas = r("var_total")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Soma_Entradas = Ent_Parcelas + Ent_Entradas
   txtTotalDinheiro = Format(Soma_Entradas, ocMONEY)
   
   'SOMAR SAIDAS
   sSQL = "SELECT ISNULL(SUM(valor), 0) AS var_saida FROM caixa_saida WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)); "
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Soma_Saidas = r("var_saida")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   txtSaida = Format(Soma_Saidas, ocMONEY)
   
   'SOMAR SALDO
   SALDO_FINAL = Soma_Entradas - Soma_Saidas
   txtSaldo = Format(SALDO_FINAL, ocMONEY)
End Sub

Private Sub cmdFecharCaixa_Click()
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   'cmdMostrar_Click
   chkCartao.Value = Checked
   Caixa_Fechamento.Show 1
End Sub

Public Sub cmdImprimir_Click()
Dim r As ADODB.Recordset
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Caixa_Fech_Imp.ReportMain1.Recordset = r

REL_Caixa_Fech_Imp.txtDHead.Caption = "FECHAMENTO DE CAIXA - DATA " & StatusBar1.Panels(5).Text
REL_Caixa_Fech_Imp.rfTroco.Caption = Format(txtTotalTroco.Text, "#,##0.00")
REL_Caixa_Fech_Imp.rfCartao.Caption = Format(txtTotalCartao.Text, "#,##0.00")
REL_Caixa_Fech_Imp.rfCheque.Caption = Format(txtTotalCheque.Text, "#,##0.00")
REL_Caixa_Fech_Imp.rfAvulso.Caption = Format(txtTotalAvulso.Text, "#,##0.00")
REL_Caixa_Fech_Imp.rfSaida.Caption = Format(txtSaida.Text, "#,##0.00")
REL_Caixa_Fech_Imp.rfDinheiro.Caption = Format(txtTotalDinheiro.Text, "#,##0.00")

If chkCartao.Value = Checked Then
   REL_Caixa_Fech_Imp.lblTotal.Caption = "SALDO C/ CARTÃO:"
   REL_Caixa_Fech_Imp.rfSaldoCCartao.Caption = Format(txtSaldo.Text, ocMONEY)
Else
   REL_Caixa_Fech_Imp.lblTotal.Caption = "SALDO S/ CARTÃO:"
   REL_Caixa_Fech_Imp.rfSaldoCCartao.Caption = Format(txtSaldo.Text, ocMONEY)
End If

'REL_Caixa_Fech_Imp.Relatorio.NomeImpressora = var_Impressora
REL_Caixa_Fech_Imp.ReportMain1.Ativar
Unload REL_Caixa_Fech_Imp

Me.Show 1
End Sub

Private Sub cmdMaqOK_Click()
   StatusBar1.Panels(2).Text = cboMaquina.Text
   cmdMostrar_Click
   frmMaquina.Visible = False
   Form_Activate
End Sub

Public Sub cmdMostrar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim SETOR_CAIXA As String
   Dim var_Setor As String
   
   If Not IsDate(mskData) Then Exit Sub
   
   'maquina
   Dim Maquina_Parcela As String
   If StatusBar1.Panels(2).Text <> "TODOS" Then
      Maquina_Parcela = "AND (parcelas.maquina = '" & StatusBar1.Panels(2).Text & "') "
   ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
      Maquina_Parcela = "AND (parcelas.maquina <> 'CAIXA') "
   End If
   
   Dim Maquina_Haver As String
   If StatusBar1.Panels(2).Text <> "TODOS" Then
      Maquina_Haver = "AND (parcelas_haver.maquina = '" & StatusBar1.Panels(2).Text & "') "
   ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
      Maquina_Haver = "AND (parcelas_haver.maquina <> 'CAIXA') "
   End If
   
   Dim Maquina_Suprimento As String
   If StatusBar1.Panels(2).Text <> "TODOS" Then
      Maquina_Suprimento = "AND (caixa_entrada.maquina = '" & StatusBar1.Panels(2).Text & "') "
   ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
      Maquina_Suprimento = "AND (caixa_entrada.maquina <> 'CAIXA') "
   End If
   
   Dim Maquina_Sangria As String
   If StatusBar1.Panels(2).Text <> "TODOS" Then
      Maquina_Sangria = "AND (caixa_saida.maquina = '" & StatusBar1.Panels(2).Text & "') "
   ElseIf StatusBar1.Panels(2).Text = "TODOS" Then
      Maquina_Sangria = "AND (caixa_saida.maquina <> 'CAIXA') "
   End If
   
   'tipo de pedido (balcao, oficina, todos)
   If StatusBar1.Panels(3).Text <> "TODOS" Then
      SETOR_CAIXA = "AND (pedidos.tipo_pedido = '" & StatusBar1.Panels(3).Text & "') "
   ElseIf StatusBar1.Panels(3).Text = "TODOS" Then
      SETOR_CAIXA = "AND (pedidos.tipo_pedido <> 'BOSTA') "
   End If
   
   'setor
   If StatusBar1.Panels(3).Text <> "TODOS" Then
      var_Setor = "AND (setor = '" & StatusBar1.Panels(3).Text & "') "
   ElseIf StatusBar1.Panels(3).Text = "TODOS" Then
      var_Setor = "AND (setor <> 'BOSTA') "
   End If
   
   'MOSTRAR As ENTRADAS
   sSQL = "SELECT " & _
      "pedidos.tipo_pedido as varTipoLanc, " & _
      "parcelas.hora AS varHora, " & _
      "parcelas.codigo AS varCodigo, " & _
      "pedidos.cod_pedido AS varCodPedido, " & _
      "cliente.nome AS varCliente, " & _
      "parcelas.forma_pgto AS varFormaPgto, " & _
      "parcelas.valor_final AS varValorLanc, " & _
      "0 AS varValorSaida, " & _
      "(parcelas.valor_final  - 0) AS campo04, " & _
      "pedidos.tipo_cartao AS varTipoCartao, " & _
      "parcelas.pagamento AS data, " & _
      "pedidos.cod_pedido AS pedido, " & _
      "pedidos.cod_cliente AS cliente, " & _
      "'' AS setor, " & _
      "parcelas.maquina " & _
      "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
      "INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
      "WHERE (parcelas.status = 1) AND (parcelas.pagamento = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & Maquina_Parcela & SETOR_CAIXA & _
      "UNION ALL "
      'MsgBox sSQL
      '
   'parcelas_haveres
   sSQL = sSQL & "SELECT " & _
      "'HAVER' as varTipoLanc, " & _
      "parcelas_haver.hora AS varHora, " & _
      "0 AS varCodigo, " & _
      "parcelas.cod_pedido AS varCodPedido, " & _
      "(SELECT nome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE (pedidos.cod_pedido = parcelas.cod_pedido)) AS varCliente, " & _
      "parcelas_haver.forma_pgto AS varFormaPgto, " & _
      "parcelas_haver.valor_haver AS varValorLanc, " & _
      "0 AS varValorSaida, " & _
      "parcelas_haver.valor_haver AS campo04, " & _
      "'' AS varTipoCartao, " & _
      "parcelas_haver.haver AS data, " & _
      "parcelas_haver.codigo AS campotc, " & _
      "0 AS cliente, " & _
      "'' AS  setor, " & _
      "'' AS maquina " & _
      "FROM parcelas_haver INNER JOIN parcelas ON parcelas_haver.cod_parcela = parcelas.codigo " & _
      "WHERE (parcelas_haver.haver = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & Maquina_Haver & _
      "UNION ALL "
   
   'suprimentos
   sSQL = sSQL & "SELECT " & _
      "'SUPRIMENTO' as varTipoLanc, " & _
      "hora AS varHora, " & _
      "0 AS varCodigo, " & _
      "0 AS varCodPedido, " & _
      "descricao AS varCliente, " & _
      "caixa_entrada.forma_pgto AS varFormaPgto, " & _
      "valor AS varValorLanc, " & _
      "0 AS varValorSaida, " & _
      "valor AS campo04, " & _
      "'' AS varTipoCartao, " & _
      "data, " & _
      "0 AS pedido, " & _
      "0 AS cliente, " & _
      "setor, " & _
      " '' AS maquina " & _
      "FROM caixa_entrada WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & Maquina_Suprimento & var_Setor & _
      "UNION ALL "
      
   'sangria
   sSQL = sSQL & "SELECT " & _
      "'SANGRIA' as varTipoLanc, " & _
      "caixa_saida.hora AS varHora, " & _
      "0 AS varCodigo, " & _
      "0 AS varCodPedido, " & _
      "caixa_saida.descricao AS varCliente, " & _
      "'DINHEIRO' AS varFormaPgto, " & _
      "0 AS varValorLanc, " & _
      "caixa_saida.valor AS varValorSaida, " & _
      "(0 - caixa_saida.valor) AS campo04, " & _
      "'' AS varTipoCartao, " & _
      "data, " & _
      "0, " & _
      "0, " & _
      "setor, " & _
      "'' AS maquina " & _
      "FROM caixa_saida WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) " & Maquina_Sangria & var_Setor & _
      " ORDER BY 2"
   
   Set r = dbData.OpenRecordset(sSQL)
   'Debug.Print sSQL
   'MsgBox r.RecordCount
   FormatarGridEntrada r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   printSQL = sSQL
   
   'MOSTRAR As ENTRADAS
   sSQL = "SELECT cliente.*, pedidos.* FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente " & _
      "WHERE (pedidos.data_compra = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (pedidos.tipo_pagamento = 'À Prazo') AND (pedidos.maquina = '" & StatusBar1.Panels(2).Text & "')"
   
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGridPrazo r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdOKData_Click()
   StatusBar1.Panels(5).Text = Format(mskData, "dd/mm/yy")
   frmData.Visible = False
   cmdMostrar_Click
   Form_Activate
End Sub

Private Sub cmdOKSetor_Click()
   StatusBar1.Panels(3).Text = cboSetor.Text
   cmdMostrar_Click
   frmSetor.Visible = False
End Sub

Private Sub cmdSalvarTroco_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim x_Troco As Long
   
   'CHECAR SE O JÁ TEM TROCO ADICIONADO PARA A DATA
   sSQL = "SELECT * FROM caixa_troco WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      dbData.Execute "UPDATE caixa_troco SET valor = " & Replace(CCur(txtTroco.Text), ",", ".") & ", maquina = '" & StatusBar1.Panels(2).Text & "' WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103));"
   Else
      x_Troco = 1
      sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_troco FROM caixa_troco;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then x_Troco = r("ultimo_troco") + 1
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      dbData.Execute "INSERT INTO caixa_troco (codigo, data, valor, maquina) VALUES (" & x_Troco & ", CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103), " & Replace(CCur(txtTroco.Text), ",", ".") & ", '" & StatusBar1.Panels(2).Text & "');"
   End If
   
   txtTroco.Text = ""
   frmTroco.Visible = False
   lblAviso1.Visible = True
   cmdMostrar_Click
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

Private Sub cmdSenha_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * FROM funcionario WHERE (senha = '" & txtSenha.Text & "') AND (nivel = 1);"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   Caixa_Fechamento.Show 1
   txtSenha.Text = ""
   frmSenha.Visible = False
Else
   ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
   txtSenha.Text = ""
   frmSenha.Visible = False
End If
End Sub

Private Sub cmdTroco_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   frmTroco.Visible = True
   lblAviso1.Visible = False
   
   sSQL = "SELECT * FROM caixa_troco WHERE (data = CONVERT(DATETIME, '" & Format(mskData, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtTroco.Text = Format(r("valor"), ocMONEY)
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   txtTroco.SetFocus
End Sub

Private Sub Form_Activate()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'MOSTRAR SE O CAIXA ESTÁ FECHADO
   sSQL = "SELECT data_abertura, maquina, status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(5).Text, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF Then
      cmdAbrirCaixa.Visible = True
      cmdFecharCaixa.Visible = False
      cmdAbrirCaixa.Caption = "Abrir Caixa"
   Else
      If CInt(ValidateNull(r("status"))) = 0 Then
         cmdFecharCaixa.Visible = True
         cmdAbrirCaixa.Visible = False
      Else
         cmdFecharCaixa.Visible = False
         cmdAbrirCaixa.Visible = True
         cmdAbrirCaixa.Caption = "Reativar"
      End If
   End If
End Sub

Private Sub Form_Load()
   Dim var_Setor As String       'mostrar o setor
   Dim var_Maquina As String     'colocar o nome da maquina na barra de status
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Setor = oIni.LerTexto("DADOS_SETOR", "setor")
   var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
   Set oIni = Nothing
   
   StatusBar1.Panels(3).Text = var_Setor
   StatusBar1.Panels(2).Text = var_Maquina

   mskData.Text = Format(Date, "dd/mm/yy")
   StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yy")
   frmTroco.Visible = False
   frmMaquina.Visible = False
   
   If StatusBar1.Panels(2).Text = "" Then StatusBar1.Panels(2).Text = "TODOS"
   cmdAbrirCaixa.Visible = False
   cmdFecharCaixa.Visible = False
   chkCartao.Value = Checked
   cmdMostrar_Click
   Set moCombo = New cComboHelper
End Sub

Private Sub FormatarGridEntradaDetalhado(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim m_Saldo As Currency
   
   With Grid
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 650
      .ColWidth(2) = 4800
      .ColWidth(3) = 1150
      .ColWidth(4) = 1150
      .ColWidth(5) = 1150
      .ColWidth(6) = 1150
      
      .TextMatrix(0, 1) = "HORA"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "TIPO"
      .TextMatrix(0, 4) = "ENTRADA"
      .TextMatrix(0, 5) = "SAÍDA"
      .TextMatrix(0, 6) = "SALDO"
      
      .Row = 0
      .Redraw = False
      
      'colocar os cabeçalho em negrito / Centralizado
      For i = 0 To .Cols - 1
         .Col = i
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      m_Saldo = 0
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("varHora"), ocHRMN)
            .TextMatrix(.Rows - 1, 2) = rTabela("varCliente")
            .TextMatrix(.Rows - 1, 3) = rTabela("var_tipo")
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("varValorLanc"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("varValorSaida"), ocMONEY)
            
            m_Saldo = m_Saldo + CCur(rTabela("varValorLanc")) - CCur(rTabela("varValorSaida"))
            .TextMatrix(.Rows - 1, 5) = Format(m_Saldo, ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'mudar a cor da coluna
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 4:   .CellBackColor = &HC0FFFF
         .Col = 5:   .CellBackColor = &HC0C0FF
      Next

      'Deixar negrito quando vencido
      For i = 1 To .Rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            If CCur(.TextMatrix(i, 4)) > 0 Then .CellFontBold = True
         Next
      Next
      
      .Redraw = True
   End With
End Sub

Private Sub FormatarGridPrazo(rTabela As ADODB.Recordset)
   Dim i  As Integer
   
   With Grid_Prazo
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 800
      .ColWidth(2) = 2000
      .ColWidth(3) = 900
      .ColWidth(4) = 950
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "NOME DO CLIENTE"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "TIPO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .Redraw = False
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("cod_pedido"), "000000")
            .TextMatrix(.Rows - 1, 2) = UCase(rTabela("nome"))
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("total"), ocMONEY)
            .TextMatrix(.Rows - 1, 4) = rTabela("pagamento")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblTotalPrazo.Caption = Format(SomaGrid(Grid_Prazo, 3), ocMONEY)
End Sub

Private Sub FormatarGridEntrada(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   Dim m_Saldo As Currency
   
   With Grid
      .Clear
      .Cols = 10
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 650
      .ColWidth(2) = 850
      .ColWidth(3) = 1300
      .ColWidth(4) = 3700
      .ColWidth(5) = 1450
      .ColWidth(6) = 1050
      .ColWidth(7) = 1050
      .ColWidth(8) = 1050
      .ColWidth(9) = 0

      
      .TextMatrix(0, 1) = "HORA"
      .TextMatrix(0, 2) = "PEDIDO"
      .TextMatrix(0, 3) = "TIPO"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "ENTRADA"
      .TextMatrix(0, 7) = "SAÍDA"
      .TextMatrix(0, 8) = "SALDO"
      .TextMatrix(0, 9) = "COD_PARCELA"
      
      .Row = 0
      
      'colocar os cabeçalho em negrito / Centralizado
      For i = 0 To .Cols - 1
         .Col = i
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      m_Saldo = 0
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = Format(rTabela("varHora"), ocHRMN)
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("varCodPedido"), "000000")
            .TextMatrix(.Rows - 1, 3) = rTabela("varTipoLanc")
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("varCliente"))
            
            If rTabela("varFormaPgto") <> "CARTAO" Then
               .TextMatrix(.Rows - 1, 5) = rTabela("varFormaPgto")
            Else
               .TextMatrix(.Rows - 1, 5) = rTabela("varFormaPgto") & " (" & rTabela("vartipocartao") & ")"
            End If
            
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("varValorLanc"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("varValorSaida"), ocMONEY)
            
            m_Saldo = m_Saldo + CCur(ValidateNull(rTabela("varValorLanc"))) - CCur(rTabela("varValorSaida"))
            .TextMatrix(.Rows - 1, 8) = Format(m_Saldo, "##,##0.00")
            .TextMatrix(.Rows - 1, 9) = Format(rTabela("varCodigo"), "000000")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'mudar a cor da coluna
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 6:   .CellBackColor = &HC0FFFF
         .Col = 7:   .CellBackColor = &HC0C0FF
      Next
      
      'Deixar negrito quando vencido
      For i = 1 To .Rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            
            If Left(.TextMatrix(i, 5), 6) = "CARTAO" Then
               .CellForeColor = &H8000&
               .CellFontBold = True
            ElseIf .TextMatrix(i, 5) = "DINHEIRO" And .TextMatrix(i, 3) <> "SANGRIA" Then
               txtTotalDinheiro.Text = Format(SomaGrid(Grid, 6), ocMONEY)
               .CellForeColor = vbBlack
            ElseIf .TextMatrix(i, 3) = "SANGRIA" And .TextMatrix(i, 5) = "DINHEIRO" Then
               .CellForeColor = vbRed
               .CellFontBold = True
            End If
         Next
      Next
      
      .Redraw = True
   End With
   
   SomaFlexDinheiro
   SomaFlexCartao
   SomaFlexCheque
   SomaFlexParcela
   SomaFlexSaida
   Mostrar_Troco
   Mostrar_Saldo
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 2)) = True Then Exit Sub
   If Grid.TextMatrix(Grid.Row, 2) = "" Or Grid.TextMatrix(Grid.Row, 3) = "" Then Exit Sub
   If Grid.TextMatrix(Grid.Row, 3) = "HAVER" Then
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 2), "OFICINA"
   Else
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 2), Grid.TextMatrix(Grid.Row, 3)
   End If
   Parcelas_Consulta_Produtos.Show 1
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

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
   Select Case Panel.Index
      Case 1
         Exit Sub
      Case 2
         frmMaquina.Visible = True
         cboMaquina.SetFocus
      Case 3
         frmSetor.Visible = True
         cboSetor.SetFocus
      Case 4
         Exit Sub
      Case 5
         frmData.Visible = True
         mskData.SetFocus
   End Select
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdSenha_Click
End Sub

Private Sub txtTroco_GotFocus()
   SelectControl txtTroco
End Sub
