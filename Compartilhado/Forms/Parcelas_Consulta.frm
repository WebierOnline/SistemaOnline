VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Parcelas_Consulta 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTA DE PARCELAS"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Parcelas_Consulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Critérios Secundários"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   120
      TabIndex        =   42
      Top             =   8340
      Width           =   4755
      Begin VB.ComboBox cboTipoCartao 
         Height          =   315
         Left            =   1500
         TabIndex        =   9
         Top             =   420
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cboParcelas 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblTipoCartao 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cartão"
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
         Left            =   1500
         TabIndex        =   44
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         Caption         =   "Parcelas"
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
         TabIndex        =   43
         Top             =   180
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   12540
      ScaleHeight     =   2085
      ScaleWidth      =   2625
      TabIndex        =   20
      Top             =   6960
      Width           =   2655
      Begin VB.CheckBox chkJuros 
         Caption         =   "Mostrar Juros"
         Height          =   195
         Left            =   1200
         TabIndex        =   34
         Top             =   1740
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas:"
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
         TabIndex        =   33
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblTotalParcelas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   32
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total:"
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
         TabIndex        =   30
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblTotalLiquido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   29
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lblTotalBruto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   28
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   27
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblTotalJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   26
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros:"
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
         Left            =   420
         TabIndex        =   25
         Top             =   600
         Width           =   525
      End
      Begin VB.Label lblTotalHaver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   24
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haveres:"
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
         TabIndex        =   23
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label lblQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1020
         TabIndex        =   22
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant.:"
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
         Left            =   300
         TabIndex        =   21
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5520
      Picture         =   "Parcelas_Consulta.frx":23D2
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   19
      Top             =   3060
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   15105
      TabIndex        =   16
      Top             =   60
      Width           =   15135
      Begin VB.Image Image1 
         Height          =   900
         Left            =   240
         Picture         =   "Parcelas_Consulta.frx":340A
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE PARCELAS"
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
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   4050
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   60
      ScaleHeight     =   2265
      ScaleWidth      =   12345
      TabIndex        =   0
      Top             =   6900
      Width           =   12375
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Impressão"
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   4920
         TabIndex        =   56
         Top             =   1440
         Width           =   5595
         Begin ChamaleonBtn.chameleonButton cmdImprime 
            Height          =   375
            Left            =   4080
            TabIndex        =   59
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Imprimir"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas_Consulta.frx":A10D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboTipoImpressao 
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   420
            Width           =   3855
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Relatório"
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
            TabIndex        =   58
            Top             =   210
            Width           =   1470
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Preencha"
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   4920
         TabIndex        =   45
         Top             =   0
         Width           =   5595
         Begin ChamaleonBtn.chameleonButton cmdCal2 
            Height          =   315
            Left            =   3360
            TabIndex        =   64
            Tag             =   "Calendario"
            Top             =   420
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            BTYPE           =   8
            TX              =   ""
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Parcelas_Consulta.frx":A129
            PICN            =   "Parcelas_Consulta.frx":A145
            PICH            =   "Parcelas_Consulta.frx":C498
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
            Left            =   2100
            TabIndex        =   63
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
            MICON           =   "Parcelas_Consulta.frx":E7EB
            PICN            =   "Parcelas_Consulta.frx":E807
            PICH            =   "Parcelas_Consulta.frx":10B5A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboTipoData 
            Height          =   315
            Left            =   3720
            TabIndex        =   60
            Top             =   420
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton OptDataIntervalo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Inter&valo"
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
            TabIndex        =   53
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton optDataUnico 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ú&nico"
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
            TabIndex        =   51
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   1020
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   1020
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCodFunc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3300
            TabIndex        =   47
            Top             =   180
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cboNome 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   420
            Visible         =   0   'False
            Width           =   5415
         End
         Begin MSMask.MaskEdBox mskFim 
            Height          =   315
            Left            =   2460
            TabIndex        =   50
            Top             =   420
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   1200
            TabIndex        =   52
            Top             =   420
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTipoData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consultar por:"
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
            Left            =   3720
            TabIndex        =   61
            Top             =   180
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblFim 
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Final"
            Enabled         =   0   'False
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
            Left            =   2460
            TabIndex        =   55
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblInicio 
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1200
            TabIndex        =   54
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
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
            TabIndex        =   49
            Top             =   800
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano"
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
            Left            =   1920
            TabIndex        =   48
            Top             =   800
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Cliente:"
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
            TabIndex        =   46
            Top             =   240
            Visible         =   0   'False
            Width           =   1470
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Critérios Principais"
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   60
         TabIndex        =   35
         Top             =   0
         Width           =   4755
         Begin VB.ComboBox cboSubCriterios 
            Height          =   315
            Left            =   3600
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   1120
         End
         Begin VB.ComboBox cboSetor 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   60
            TabIndex        =   1
            Top             =   460
            Width           =   1095
         End
         Begin VB.ComboBox cboCriterios 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2310
            TabIndex        =   3
            Top             =   460
            Width           =   1275
         End
         Begin VB.ComboBox cboSituacao 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1150
            TabIndex        =   2
            Top             =   460
            Width           =   1155
         End
         Begin VB.ComboBox cboForma 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1200
            TabIndex        =   6
            Top             =   1020
            Width           =   1395
         End
         Begin VB.ComboBox cboOrdem 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   1020
            Width           =   1095
         End
         Begin VB.ComboBox cboTipo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2640
            TabIndex        =   7
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label lblSubCriterios 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Critérios"
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
            Left            =   3600
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Setor:"
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
            TabIndex        =   41
            Top             =   240
            Width           =   525
         End
         Begin VB.Label lblCriterios 
            AutoSize        =   -1  'True
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
            Left            =   2310
            TabIndex        =   40
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Left            =   1150
            TabIndex        =   39
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Venda:"
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
            TabIndex        =   38
            Top             =   795
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            Left            =   60
            TabIndex        =   37
            Top             =   795
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pgto"
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
            Left            =   2640
            TabIndex        =   36
            Top             =   795
            Width           =   1245
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   615
         Left            =   10620
         TabIndex        =   13
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
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
         MICON           =   "Parcelas_Consulta.frx":12EAD
         PICN            =   "Parcelas_Consulta.frx":12EC9
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
         Left            =   10620
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
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
         MICON           =   "Parcelas_Consulta.frx":137A3
         PICN            =   "Parcelas_Consulta.frx":137BF
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
         Height          =   615
         Left            =   10620
         TabIndex        =   15
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Imprimir Agrupados"
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
         MICON           =   "Parcelas_Consulta.frx":13AD9
         PICN            =   "Parcelas_Consulta.frx":13AF5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         X1              =   4620
         X2              =   4620
         Y1              =   60
         Y2              =   1920
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5775
      Left            =   60
      TabIndex        =   18
      Top             =   1080
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   10186
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   31
      Top             =   9195
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22595
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23:22"
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
End
Attribute VB_Name = "Parcelas_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private printSQL As String
Dim vImpressoraNormal As String
Dim oIni As Ini
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vNome, vCelular, vEndereco, vNum, vBairro, vCidade, vUF, vReferencia As String

Private Sub Limpar_Grid()
   Dim i As Integer
   
   With Grid
      .Visible = False
      
      .Clear
      .Cols = 15
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 900
      .ColWidth(2) = 0
      .ColWidth(3) = 4200
      .ColWidth(4) = 0
      .ColWidth(5) = 850
      .ColWidth(6) = 1250
      .ColWidth(7) = 900
      .ColWidth(8) = 1140
      .ColWidth(9) = 950
      .ColWidth(10) = 950
      .ColWidth(11) = 950
      .ColWidth(12) = 950
      .ColWidth(13) = 950
      .ColWidth(14) = 900
      
      .TextMatrix(0, 1) = "ORIGEM"
      .TextMatrix(0, 2) = "CÓD."
      .TextMatrix(0, 3) = "CLIENTE"
      .TextMatrix(0, 4) = "CELULAR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "TIPO"
      .TextMatrix(0, 7) = "VENC"
      .TextMatrix(0, 8) = "SUBTOTAL"
      .TextMatrix(0, 9) = "HAVER"
      .TextMatrix(0, 10) = "ATRAZO"
      .TextMatrix(0, 11) = "JUROS"
      .TextMatrix(0, 12) = "TOTAL"
      .TextMatrix(0, 13) = "STATUS"
      .TextMatrix(0, 14) = "PGTO"
      
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
      .Redraw = False
      .rows = .rows + 1
      
      i = i + 1
      .rows = .rows - 1
      
      .Redraw = True
      .Visible = True
      
      'somar a colunas totais
   'somar a colunas totais
   lblQuant.Caption = Grid.rows - 1
   lblTotalParcelas.Caption = Format(0, ocMONEY)
   lblTotalBruto.Caption = Format(0, ocMONEY)
   lblTotalHaver.Caption = Format(0, ocMONEY)
   lblTotalJuros.Caption = Format(0, ocMONEY)
   lblTotalLiquido.Caption = Format(0, ocMONEY)
   End With
End Sub

Private Sub Preencher_TipoDatas()
   cboTipoData.Clear
   cboTipoData.AddItem "PGTO"
   cboTipoData.AddItem "VENC."
End Sub

Private Sub Preencher_Forma()
cboForma.Clear
cboForma.AddItem "TODAS"
cboForma.AddItem "À VISTA"
cboForma.AddItem "À PRAZO"
End Sub

Private Sub Preencher_Ordem()
   cboOrdem.Clear
   cboOrdem.AddItem "NOME"
   cboOrdem.AddItem "TIPO"
   cboOrdem.AddItem "PGTO"
   cboOrdem.AddItem "VENC."
   cboOrdem.AddItem "VALOR"
End Sub

Private Sub Preencher_Setor()
   cboSetor.Clear
   cboSetor.AddItem "TODOS"
   cboSetor.AddItem "VENDAS"
   cboSetor.AddItem "OFICINA"
   cboSetor.AddItem "RECEBER"
   cboSetor.AddItem "ALUGUEL"
End Sub

Private Sub Preencher_Situacao()
   cboSituacao.Clear
   cboSituacao.AddItem "TODAS"
   cboSituacao.AddItem "PAGAS"
   cboSituacao.AddItem "À PAGAR"
   cboSituacao.AddItem "VENCIDAS"
End Sub

Private Sub Preencher_SubCriterios()
   cboSubCriterios.Clear
   cboSubCriterios.AddItem "NENHUM"
   cboSubCriterios.AddItem "MENSAL"
   'cboSubCriterios.AddItem "À PRAZO"
End Sub

Private Sub Preencher_Tipo()
   cboTipo.Clear
   cboTipo.AddItem "TODOS"
   cboTipo.AddItem "DINHEIRO"
   cboTipo.AddItem "CARTÃO"
   cboTipo.AddItem "PIX"
   cboTipo.AddItem "PROMISSÓRIA"
   cboTipo.AddItem "CHEQUE"
   cboTipo.AddItem "SEM CARTÃO"
End Sub

Private Sub Preencher_TipoImpressao()
cboTipoImpressao.Clear

If cboCriterios.Text = "TODOS" Then
    cboTipoImpressao.AddItem "RELATÓRIO NORMAL"
    cboTipoImpressao.AddItem "RELATÓRIO AGRUPADO"
ElseIf cboCriterios.Text = "DATA" Then
    cboTipoImpressao.AddItem "RELATÓRIO NORMAL"
    'cboTipoImpressao.AddItem "RELATÓRIO AGRUPADO"
ElseIf cboCriterios.Text = "MENSAL" Then
    cboTipoImpressao.AddItem "RELATÓRIO NORMAL"
    'cboTipoImpressao.AddItem "RELATÓRIO AGRUPADO"
ElseIf cboCriterios.Text = "CLIENTE" Then
    cboTipoImpressao.AddItem "RELATÓRIO NORMAL"
    cboTipoImpressao.AddItem "RELATÓRIO UNIFICADO"
ElseIf cboCriterios.Text = "VENDEDOR" Then
    cboTipoImpressao.AddItem "RELATÓRIO NORMAL"
    'cboTipoImpressao.AddItem "RELATÓRIO AGRUPADO"
    'cboTipoImpressao.AddItem "RELATÓRIO UNIFICADO"
End If
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
End Sub

Private Sub cboCriterios_Change()
   If cboCriterios.Text = "TODOS" Then
      lblSubCriterios.Visible = False
      cboSubCriterios.Visible = False
      
      'lblDatas.Visible = False
      'cboDatas.Visible = False
      
      lblTipoData.Visible = False
      cboTipoData.Visible = False
      
      lblNome.Visible = False
      cboNome.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      optDataUnico.Visible = False
      OptDataIntervalo.Visible = False
      lblInicio.Visible = False
      lblFim.Visible = False
      mskInicio.Visible = False
      mskFim.Visible = False
      cmdCal1.Visible = False
      cmdCal2.Visible = False
   
   ElseIf cboCriterios.Text = "DATA" Then
      lblSubCriterios.Visible = False
      cboSubCriterios.Visible = False
      
      'lblDatas.Visible = True
      'cboDatas.Visible = True
      
      lblTipoData.Visible = True
      cboTipoData.Visible = True
      
      cboTipoData.Top = 420
      cboTipoData.Left = 3720
      
      lblTipoData.Top = 180
      lblTipoData.Left = 3720
      
      lblNome.Visible = False
      cboNome.Visible = False
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      optDataUnico.Visible = True
      OptDataIntervalo.Visible = True
      lblInicio.Visible = True
      lblFim.Visible = True
      mskInicio.Visible = True
      mskFim.Visible = True
      cmdCal1.Visible = True
      cmdCal2.Visible = True
   
   ElseIf cboCriterios.Text = "MENSAL" Then
      lblSubCriterios.Visible = False
      cboSubCriterios.Visible = False
      'lblDatas.Visible = True
      'cboDatas.Visible = True
      
      lblTipoData.Visible = True
      cboTipoData.Visible = True
      
      cboTipoData.Top = 1020
      cboTipoData.Left = 3300
      
      lblTipoData.Top = 780
      lblTipoData.Left = 3300
      
      lblNome.Visible = False
      cboNome.Visible = False
      
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      
      optDataUnico.Visible = False
      OptDataIntervalo.Visible = False
      lblInicio.Visible = False
      lblFim.Visible = False
      mskInicio.Visible = False
      mskFim.Visible = False
      cmdCal1.Visible = False
      cmdCal2.Visible = False
   
   ElseIf cboCriterios.Text = "CLIENTE" Then
      lblNome.Caption = "Nome do Cliente"
      lblSubCriterios.Visible = True
      cboSubCriterios.Visible = True
      
      'lblDatas.Visible = False
      'cboDatas.Visible = False
      
      lblTipoData.Visible = False
      cboTipoData.Visible = False
      
      lblNome.Visible = True
      cboNome.Visible = True
      cboNome.Clear
      cboNome.Text = ""
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      optDataUnico.Visible = False
      OptDataIntervalo.Visible = False
      lblInicio.Visible = False
      lblFim.Visible = False
      mskInicio.Visible = False
      mskFim.Visible = False
      cmdCal1.Visible = False
      cmdCal2.Visible = False
      cboSubCriterios.ListIndex = 0
   
   ElseIf cboCriterios.Text = "VENDEDOR" Then
      lblNome.Caption = "Nome do Vendedor"
      lblSubCriterios.Visible = True
      cboSubCriterios.Visible = True
      
      'lblDatas.Visible = False
      'cboDatas.Visible = False
      
      lblTipoData.Visible = False
      cboTipoData.Visible = False
      
      lblNome.Visible = True
      cboNome.Visible = True
      cboNome.Clear
      cboNome.Text = ""
      
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      
      optDataUnico.Visible = False
      OptDataIntervalo.Visible = False
      lblInicio.Visible = False
      lblFim.Visible = False
      mskInicio.Visible = False
      mskFim.Visible = False
      cmdCal1.Visible = False
      cmdCal2.Visible = False
      cboSubCriterios.ListIndex = 0
   
   End If
   
   If cboCriterios.Text <> "" Then cmdExibir_Click
   If cboCriterios.Text <> "" Then cboCriterios_LostFocus
End Sub

Private Sub cboCriterios_Click()
   cboCriterios_Change
End Sub

Private Sub cboCriterios_GotFocus()
   Preencher_Criterio
   moCombo.AttachTo cboCriterios
End Sub

Private Sub Preencher_TipoCartao()
   cboTipoCartao.Clear
   cboTipoCartao.AddItem "TODOS"
   cboTipoCartao.AddItem "DEBITO"
   cboTipoCartao.AddItem "CREDITO"
End Sub

Private Sub Preencher_Parcelas()
   cboParcelas.Clear
   cboParcelas.AddItem "SÓ ENTRADA"
   cboParcelas.AddItem "TODAS"
End Sub

Private Sub Preencher_Criterio()
   cboCriterios.Clear
   cboCriterios.AddItem "TODOS"
   cboCriterios.AddItem "DATA"
   cboCriterios.AddItem "MENSAL"
   cboCriterios.AddItem "CLIENTE"
   cboCriterios.AddItem "VENDEDOR"
End Sub

Private Sub cboCriterios_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Private Sub cboCriterios_LostFocus()
cboTipoImpressao.Clear
cboTipoImpressao.Text = ""
Preencher_TipoImpressao
If cboCriterios.Text <> "" Then cboTipoImpressao.ListIndex = 0
End Sub

Private Sub cboForma_Change()
   If cboForma.Text = "À PRAZO" Then
      lblParcelas.Visible = True
      cboParcelas.Visible = True
   Else
      lblParcelas.Visible = False
      cboParcelas.Visible = False
   End If
   
   If cboForma.Text = "À VISTA" Then
      cboSituacao.ListIndex = 1
   End If
   
   If cboForma.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboForma_Click()
   cboForma_Change
End Sub

Private Sub cboForma_GotFocus()
Preencher_Forma
moCombo.AttachTo cboForma
End Sub

Private Sub cboForma_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMes.Clear
   For vMes = 1 To 12
      cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMes
End Sub

Private Sub cboMes_LostFocus()
   cboAno.SetFocus
End Sub

Private Sub cboNome_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboNome.Clear
   
   If cboCriterios.Text = "CLIENTE" Then
      sSQL = "SELECT * FROM cliente ORDER BY nome;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboNome.AddItem r("nome")
         cboNome.ItemData(cboNome.NewIndex) = r("codigo")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
   ElseIf cboCriterios.Text = "VENDEDOR" Then
      sSQL = "SELECT * FROM funcionario ORDER BY nome;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboNome.AddItem r("nome")
         cboNome.ItemData(cboNome.NewIndex) = r("codigo")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
   
   moCombo.AttachTo cboNome
End Sub

Private Sub cboNome_LostFocus()
   If cboNome.Text = "" Then txtCodFunc.Text = "": Exit Sub
   If cboNome.ListIndex = -1 Then txtCodFunc.Text = "": Exit Sub
   txtCodFunc = cboNome.ItemData(cboNome.ListIndex)
   Exit Sub
End Sub

Private Sub cboOrdem_Change()
   If cboOrdem.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboOrdem_Click()
cboOrdem_Change
End Sub


Private Sub cboOrdem_GotFocus()
   Preencher_Ordem
   moCombo.AttachTo cboOrdem
End Sub

Private Sub cboOrdem_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboParcelas_Change()
   If cboParcelas.Text = "SÓ ENTRADA" Then
      cboSituacao.ListIndex = 1
   End If
   
   If cboParcelas.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboParcelas_Click()
   cboParcelas_Change
End Sub

Private Sub cboParcelas_GotFocus()
   Preencher_Parcelas
   moCombo.AttachTo cboParcelas
End Sub

Private Sub cboParcelas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSetor_Change()
   If cboSetor.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboSetor_Click()
   cboSetor_Change
End Sub

Private Sub cboSetor_GotFocus()
   Preencher_Setor
   moCombo.AttachTo cboSetor
End Sub

Private Sub cboSETOR_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSituacao_Change()
   If cboSituacao.Text <> "PAGAS" Then
      cboParcelas.ListIndex = 1
   End If
   
   If cboSituacao.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboSituacao_Click()
   cboSituacao_Change
End Sub

Private Sub cboSituacao_GotFocus()
   Preencher_Situacao
   moCombo.AttachTo cboSituacao
End Sub

Private Sub cboSituacao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSubCriterios_Change()
If cboCriterios.Text = "CLIENTE" Or cboCriterios.Text = "VENDEDOR" Then
   If cboSubCriterios.Text = "MENSAL" Then
      lblMes.Visible = True
      cboMes.Visible = True
      lblAno.Visible = True
      cboAno.Visible = True
      lblTipoData.Visible = True
      cboTipoData.Visible = True
      cboTipoData.Top = 1020
      cboTipoData.Left = 3300
      lblTipoData.Top = 780
      lblTipoData.Left = 3300
   ElseIf cboSubCriterios.Text = "NENHUM" Then
      lblMes.Visible = False
      cboMes.Visible = False
      lblAno.Visible = False
      cboAno.Visible = False
      lblTipoData.Visible = False
      cboTipoData.Visible = False
   End If
End If
If cboSubCriterios.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboSubCriterios_Click()
   cboSubCriterios_Change
End Sub

Private Sub cboSubCriterios_GotFocus()
   Preencher_SubCriterios
   moCombo.AttachTo cboSubCriterios
End Sub

Private Sub cboSubCriterios_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipo_Change()
   If cboTipo.Text = "CARTÃO" Then
      lblTipoCartao.Visible = True
      cboTipoCartao.Visible = True
   Else
      lblTipoCartao.Visible = False
      cboTipoCartao.Visible = False
   End If
   
   If cboTipo.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboTipo_Click()
   cboTipo_Change
End Sub

Private Sub cboTipo_GotFocus()
   Preencher_Tipo
   moCombo.AttachTo cboTipo
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipoCartao_Change()
   If cboTipoCartao.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboTipoCartao_GotFocus()
   Preencher_TipoCartao
   moCombo.AttachTo cboTipoCartao
End Sub

Private Sub cboTipoCartao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipoData_Change()
If cboTipoData.Text <> "" Then cmdExibir_Click
End Sub

Private Sub cboTipoData_GotFocus()
   Preencher_TipoDatas
   moCombo.AttachTo cboTipoData
End Sub


Private Sub cboTipoData_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipoImpressao_GotFocus()
Preencher_TipoImpressao
End Sub


Private Sub chameleonButton1_Click()

sSQL = "SELECT DISTINCT cliente.CODIGO AS varClienteCod, cliente.Nome, CASE WHEN cliente.telefone1 IS NULL THEN cliente.celular WHEN cliente.telefone1 = '' THEN cliente.celular ELSE cliente.telefone1 END AS vContato, " & _
       "(SELECT ISNULL(SUM(parcelas.VALOR), 0) AS Expr1 FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO WHERE (pedidos.COD_CLIENTE = cliente.CODIGO) AND (parcelas.STATUS = 0) AND (parcelas.DATA < GETDATE())) AS varSomaParcelasVencidas, " & _
       "(SELECT ISNULL(SUM(parcelas_2.VALOR), 0) AS Expr1 FROM parcelas AS parcelas_2 INNER JOIN pedidos AS pedidos_2 ON parcelas_2.COD_PEDIDO = pedidos_2.COD_PEDIDO WHERE (pedidos_2.COD_CLIENTE = cliente.CODIGO) AND (parcelas_2.STATUS = 0) AND (parcelas_2.DATA > GETDATE())) AS varSomaParcelasAVencer " & _
       "FROM parcelas AS parcelas_1 INNER JOIN pedidos AS pedidos_1 ON parcelas_1.COD_PEDIDO = pedidos_1.COD_PEDIDO INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO " & _
       "Where (parcelas_1.Status = 0) " & _
       "GROUP BY cliente.CODIGO, cliente.Nome, cliente.TELEFONE1, cliente.CELULAR " & _
       "ORDER BY cliente.Nome"

Set r = dbData.OpenRecordset(sSQL)

'Debug.Print sSQL
printSQL = sSQL

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Parcelas_Agrupado.Relatorio.Recordset = r

'REL_Parcelas_Agrupado.rfQuant.Caption = lblQuant.Caption
'REL_Parcelas_Agrupado.rfsubtotal.Caption = Format(lblTotalBruto.Caption, "##,##0.00")

'REL_Parcelas_Agrupado.rfForma.Caption = cboForma.Text
'REL_Parcelas_Agrupado.rfTipo.Caption = cboTipo.Text

REL_Parcelas_Agrupado.Relatorio.NomeImpressora = vImpressoraNormal
REL_Parcelas_Agrupado.Relatorio.Ativar
Unload REL_Parcelas_Agrupado

Me.Show 1
End Sub

Private Sub chkJuros_Click()
cmdExibir_Click
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

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
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

mskFim = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdExibir_Click()
   Dim RESULTADO As Currency
   Dim INDICE As String    'INDICE
   Dim Tipo As String      'STATUS
   Dim DATAS As String     'INDICE DAS DATAS
   
   Dim var_Forma As String       'FORMA DE PGTO
   Dim varSituacaoPGTO As String    'TIPO DE PGTO
   Dim tipo_Cartao As String     'TIPO DE CARTÃO
   Dim var_JurosDia As Double
   Dim var_TipoPedido As String
   
   Dim oCfg As ConfigItem
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim var_tipojuros As Integer
   Dim vValorCobrarJuros As String
   
   If cboCriterios.Text = "" Then Exit Sub
   If cboCriterios.Text = "TODOS" Then
        If cboSituacao.Text <> "VENCIDAS" And cboSituacao.Text <> "À PAGAR" Then
               MsgBox "A consulta resultará no volume de informação muito grande. Reformule a consulta!", vbInformation, "Aviso do Sistema"
               Limpar_Grid
               Exit Sub
        End If
    End If

   'Verifica o tipo de pedido
   If cboSetor.Text = "TODOS" Then
      var_TipoPedido = "(tipo_pedido <> '')"
   ElseIf cboSetor.Text = "ORÇAMENTO" Then
      var_TipoPedido = "(tipo_pedido = 'ORÇAMENTO')"
   ElseIf cboSetor.Text = "VENDAS" Then
      var_TipoPedido = "(tipo_pedido = 'VENDA')"
   ElseIf cboSetor.Text = "OFICINA" Then
      var_TipoPedido = "(tipo_pedido = 'OFICINA')"
   ElseIf cboSetor.Text = "RECEBER" Then
      var_TipoPedido = "(tipo_pedido = 'RECEBER')"
   ElseIf cboSetor.Text = "ALUGUEL" Then
      var_TipoPedido = "(tipo_pedido = 'ALUGUEL')"
   End If

   'Ver o valor do juros
   Set oCfg = sysConfig("JUROS_DIA")
   var_JurosDia = CCur(oCfg.Value)
   Set oCfg = Nothing
   
   'tipo de juros
    Set oCfg = sysConfig("TIPO_JUROS")
   var_tipojuros = oCfg.Value
   Set oCfg = Nothing
   
   'Tipo de juros (parcial ou total)
   If var_tipojuros = 0 Then              'Juros sobre o saldo restante
      vValorCobrarJuros = "(parcelas.valor - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo)))"
   ElseIf var_tipojuros = 1 Then          'Juros sobre o valor da parcela
      vValorCobrarJuros = "parcelas.valor"
   End If
      
   If chkJuros.Value = Checked Then
        sSQL = "SELECT pedidos.tipo_pedido AS var_tipo, parcelas.cod_pedido as var_codigo, cliente.nome, CASE WHEN cliente.telefone1 IS NULL THEN cliente.celular WHEN cliente.telefone1 = '' THEN cliente.celular ELSE cliente.telefone1 END AS celular, pedidos.tipo_pagamento, ISNULL(parcelas.FORMA_PGTO, '') AS var_FormaPgto, parcelas.status,  parcelas.data, parcelas.pagamento AS var_DataPgto, parcelas.valor_final, " & _
           "parcelas.valor as varValorParc, " & _
           "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (parcelas_haver.cod_parcela = parcelas.codigo)) AS var_haver, " & _
           "CASE parcelas.status WHEN 0 THEN (CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END) ELSE (parcelas.dias_atrazo) END AS varDiasJuros, " & _
           "CASE parcelas.status WHEN 0 THEN ((((" & vValorCobrarJuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END)) ELSE (parcelas.juros) END AS varValorJuros, " & _
           "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS varSituacaoPGTO, " & _
           "(parcelas.valor - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) + CASE parcelas.status WHEN 0 THEN ((((" & vValorCobrarJuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, GETDATE(), parcelas.data) ELSE 0 END)) ELSE (parcelas.juros) END AS var_Total, " & _
           "(CASE parcelas.status WHEN 0 THEN ((((" & vValorCobrarJuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END)) ELSE (parcelas.juros) END) + parcelas.valor as varParcComJuros,  " & _
           "(((CASE parcelas.status WHEN 0 THEN ((((" & vValorCobrarJuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, parcelas.data, GETDATE()) ELSE 0 END)) ELSE (parcelas.juros) END) + parcelas.valor) - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (parcelas_haver.cod_parcela = parcelas.codigo))) AS varTotalLiquido " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "INNER JOIN cliente ON cliente.codigo = pedidos.cod_cliente " & _
           "WHERE " & var_TipoPedido & " AND "
    Else
        sSQL = "SELECT pedidos.tipo_pedido AS var_tipo, parcelas.cod_pedido as var_codigo, cliente.nome, CASE WHEN cliente.telefone1 IS NULL THEN cliente.celular WHEN cliente.telefone1 = '' THEN cliente.celular ELSE cliente.telefone1 END AS celular, pedidos.tipo_pagamento, ISNULL(parcelas.FORMA_PGTO, '') AS var_FormaPgto, parcelas.status,  parcelas.data, parcelas.pagamento AS var_DataPgto, parcelas.valor_final, " & _
           "parcelas.valor as varValorParc, " & _
           "(SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (parcelas_haver.cod_parcela = parcelas.codigo)) AS var_haver, " & _
           "'0' AS varDiasJuros, " & _
           "'0' AS varValorJuros, " & _
           "CASE parcelas.status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS varSituacaoPGTO, " & _
           "(parcelas.valor - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (cod_parcela = parcelas.codigo))) + CASE parcelas.status WHEN 0 THEN ((((" & vValorCobrarJuros & " * " & Replace(var_JurosDia, ",", ".") & ") / 100) * CASE WHEN parcelas.data <= GETDATE() THEN DATEDIFF(day, GETDATE(), parcelas.data) ELSE 0 END)) ELSE (parcelas.juros) END AS var_Total, " & _
           "parcelas.valor as varParcComJuros,  " & _
           "(parcelas.valor - (SELECT ISNULL(SUM(valor_haver), 0) FROM parcelas_haver WHERE (parcelas_haver.cod_parcela = parcelas.codigo))) AS varTotalLiquido " & _
           "FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
           "INNER JOIN cliente ON cliente.codigo = pedidos.cod_cliente " & _
           "WHERE " & var_TipoPedido & " AND "
    End If
      
   
   'indice
   If cboOrdem.Text = "NOME" Then
      INDICE = "nome "
   ElseIf cboOrdem.Text = "VENC." Then
      INDICE = "parcelas.data "
   ElseIf cboOrdem.Text = "PGTO" Then
      INDICE = "parcelas.pagamento "
   ElseIf cboOrdem.Text = "TIPO" Then
      INDICE = "pedidos.pagamento "
   ElseIf cboOrdem.Text = "VALOR" Then
      INDICE = "var_Total DESC"
   Else
      INDICE = "parcelas.data "
   End If
   
   
   If cboSituacao.Text = "TODAS" Then
      Tipo = ""
   ElseIf cboSituacao.Text = "PAGAS" Then
      Tipo = " AND (parcelas.status = 1) "
   ElseIf cboSituacao.Text = "À PAGAR" Then
      Tipo = " AND (parcelas.status = 0) "
   ElseIf cboSituacao.Text = "VENCIDAS" Then
      Tipo = " AND (parcelas.status = 0) AND (parcelas.data < CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103)) "
   End If
   
   
   'If cboDatas.Text = "PGTO" Then
   '   DATAS = "parcelas.pagamento "
  ' ElseIf cboDatas.Text = "VENC." Then
   '   DATAS = "parcelas.data "
  ' Else
   '   DATAS = "parcelas.pagamento "
  ' End If
   
   If cboTipoData.Text = "PGTO" Then
      DATAS = "parcelas.pagamento "
   ElseIf cboTipoData.Text = "VENC." Then
      DATAS = "parcelas.data "
   Else
      DATAS = "parcelas.pagamento "
   End If
   
   
   If cboForma.Text = "TODAS" Then
      var_Forma = ""
   ElseIf cboForma.Text = "À VISTA" Then
      var_Forma = " AND (pedidos.tipo_pagamento = 'À Vista') "
   ElseIf cboForma.Text = "À PRAZO" Then
      var_Forma = " AND (pedidos.tipo_pagamento = 'À Prazo') "
   End If
   
   
'   If cboTipo.Text = "TODOS" Then
'      varSituacaoPGTO = ""
'   ElseIf cboTipo.Text = "AVULSO" Then
'      varSituacaoPGTO = " AND (pedidos.pagamento = 'AVULSO') "
'   ElseIf cboTipo.Text = "CARTÃO" Then
'      varSituacaoPGTO = " AND (pedidos.pagamento = 'CARTAO') "
'   ElseIf cboTipo.Text = "CHEQUE" Then
'      varSituacaoPGTO = " AND (pedidos.pagamento = 'CHEQUE') "
'   ElseIf cboTipo.Text = "DINHEIRO" Then
'      varSituacaoPGTO = " AND (pedidos.pagamento = 'DINHEIRO') "
'   ElseIf cboTipo.Text = "PROMISSÓRIA" Then
'      varSituacaoPGTO = " AND (pedidos.pagamento = 'PROMISSORIA') "
'   ElseIf cboTipo.Text = "SEM CARTÃO" Then
'      varSituacaoPGTO = " AND (pedidos.pagamento <> 'CARTAO') "
'   End If

   If cboTipo.Text = "TODOS" Then
      varSituacaoPGTO = ""
   ElseIf cboTipo.Text = "AVULSO" Then
      varSituacaoPGTO = " AND (PARCELAS.FORMA_PGTO = 'AVULSO') "
   ElseIf cboTipo.Text = "CARTÃO" Then
      varSituacaoPGTO = " AND (PARCELAS.FORMA_PGTO = 'CARTAO') "
   ElseIf cboTipo.Text = "CHEQUE" Then
      varSituacaoPGTO = " AND (PARCELAS.FORMA_PGTO = 'CHEQUE') "
   ElseIf cboTipo.Text = "DINHEIRO" Then
      varSituacaoPGTO = " AND (PARCELAS.FORMA_PGTO = 'DINHEIRO') "
   ElseIf cboTipo.Text = "PROMISSÓRIA" Then
      varSituacaoPGTO = " AND (PARCELAS.FORMA_PGTO = 'PROMISSORIA') "
   ElseIf cboTipo.Text = "SEM CARTÃO" Then
      varSituacaoPGTO = " AND (PARCELAS.FORMA_PGTO <> 'CARTAO') "
   End If
   
   
   If cboTipo.Text = "CARTÃO" Then
      If cboTipoCartao.Text = "TODOS" Then
         tipo_Cartao = ""
      ElseIf cboTipoCartao.Text = "DEBITO" Then
         tipo_Cartao = " AND (PARCELAS.tipo_cartao = 'D') "
      ElseIf cboTipoCartao.Text = "CREDITO" Then
         tipo_Cartao = " AND (PARCELAS.tipo_cartao = 'C') "
      End If
   Else
      tipo_Cartao = ""
   End If
   
   
   If cboCriterios.Text = "TODOS" Then
      sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
         Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
   
   ElseIf cboCriterios.Text = "DATA" Then
      If optDataUnico.Value = True Then
         If Not IsDate(mskInicio) Then Exit Sub
         sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
            "AND (" & DATAS & " = CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103)) " & Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
        
      ElseIf OptDataIntervalo.Value = True Then
         If Not IsDate(mskInicio) And Not IsDate(mskFim) Then Exit Sub
         sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
            " AND (" & DATAS & " >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (" & DATAS & " <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
            Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY  " & INDICE
      End If
      
   ElseIf cboCriterios.Text = "MENSAL" Then
      If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
      sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
         "AND (MONTH(" & DATAS & ") = " & cboMes.ListIndex + 1 & ") AND (YEAR(" & DATAS & ") = " & cboAno & ") " & Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
      
   ElseIf cboCriterios.Text = "CLIENTE" Then
      If cboNome.Text = "" Then Exit Sub
      If cboSubCriterios.Text = "MENSAL" Then
         If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
         sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
            "AND (MONTH(parcelas.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(parcelas.data) = " & cboAno & ") AND (nome = '" & cboNome.Text & "') " & _
            Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
      Else
         sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
            "AND (nome = '" & cboNome.Text & "') " & Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
      End If
      
   ElseIf cboCriterios.Text = "VENDEDOR" Then
      If cboNome.Text = "" Then Exit Sub
      If cboSubCriterios.Text = "MENSAL" Then
         If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
         sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
            "AND (MONTH(" & DATAS & ") = " & cboMes.ListIndex + 1 & ") AND (YEAR(" & DATAS & ") = " & cboAno & ") " & _
            "AND (pedidos.cod_funcionario = " & txtCodFunc.Text & ") " & Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
      Else
         sSQL = sSQL & IIf(cboParcelas.Text = "SÓ ENTRADA", "(pedidos.entrada <> 0) AND (parcelas.numero = 1)", "(pedidos.cod_pedido > 0) ") & _
            "AND (pedidos.cod_funcionario = " & txtCodFunc.Text & ") " & Tipo & tipo_Cartao & var_Forma & varSituacaoPGTO & " ORDER BY " & INDICE
      End If
      
   Else
      sSQL = sSQL & " WHERE false"
   
   End If
   
   'Debug.Print sSQL
   
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGrid r, var_JurosDia
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   printSQL = sSQL
   picAguarde.Visible = False
End Sub

Private Sub FormatarGrid2(rTabela As ADODB.Recordset, ByVal percJuros As Currency)
   Dim i As Integer, j As Integer
   
   'Dim vJuros As Currency
   'Dim vHaver As Currency
   'Dim vTotal As Currency
   
   picAguarde.Visible = True
   DoEvents
   
   With Grid
      .Visible = False
      
      .Clear
      .Cols = 16
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 800
      .ColWidth(2) = 0
      .ColWidth(3) = 4200
      .ColWidth(4) = 0
      .ColWidth(5) = 800
      .ColWidth(6) = 1250
      .ColWidth(7) = 900
      .ColWidth(8) = 900
      .ColWidth(9) = 900
      .ColWidth(10) = 900
      .ColWidth(11) = 650
      .ColWidth(12) = 900
      .ColWidth(13) = 900
      .ColWidth(14) = 950
      .ColWidth(15) = 850
      
      .TextMatrix(0, 1) = "Origem"    'pedidos.tipo_pedido
      .TextMatrix(0, 2) = "Cód."      'PARCELAS.COD_PEDIDO
      .TextMatrix(0, 3) = "Cliente"   'cliente.nome
      .TextMatrix(0, 4) = "Celular"   'cliente.celular
      .TextMatrix(0, 5) = "Forma"     'pedidos.tipo_pagamento
      .TextMatrix(0, 6) = "Tipo"      'pedidos.pagamento
      .TextMatrix(0, 7) = "Venc"      'parcelas.data
      .TextMatrix(0, 8) = "Subtotal"  'varValorParc
      .TextMatrix(0, 9) = "Haver"     'se parcelas.haver = 1 entao soma(parcelas_haver.valor_haver where parcelas.codigo = parcelas_haver.cod_parcelas
      .TextMatrix(0, 10) = "Total"   'parcelas.dias_atrazo
      .TextMatrix(0, 11) = "Dias"   'parcelas.dias_atrazo
      .TextMatrix(0, 12) = "Juros"    'parcelas.juros
      .TextMatrix(0, 13) = "Deve"    'parcelas.valor_final
      .TextMatrix(0, 14) = "Status"   'se parcelas.status = 1 entao texto 'pago'
      .TextMatrix(0, 15) = "Pgto"     'var_DataPgto
      
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
      
      .Redraw = False
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'Cálcula os juros e o total da parcela
            'vJuros = 0
            'vHaver = 0
            
            'If CInt(rTabela("status")) = 0 Then
             '  If cboJuros.Text = "SIM" Then vJuros = CCur(rTabela("valor")) * (percJuros / 100)
             '  vHaver = rTabela("var_haver")
            'Else
            '   vJuros = rTabela("varvalorjuros")
            '   vHaver = 0
            'End If
            
            'vTotal = CCur(rTabela("valor")) - vHaver + vJuros
            
            .TextMatrix(.rows - 1, 1) = rTabela("var_tipo")
            .TextMatrix(.rows - 1, 2) = rTabela("var_codigo")
            .TextMatrix(.rows - 1, 3) = rTabela("nome") & IIf(Trim(ValidateNull(rTabela("celular"))) = "", "", "     (" & Right$(ValidateNull(rTabela("celular")), 9) & ")")
            
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("celular"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("tipo_pagamento"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_FormaPgto"))
            .TextMatrix(.rows - 1, 7) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 8) = Format(rTabela("varValorParc"), ocMONEY)
            .TextMatrix(.rows - 1, 9) = Format(rTabela("var_haver"), ocMONEY)
            .TextMatrix(.rows - 1, 10) = Format(rTabela("varValorComHaver"), ocMONEY)
            .TextMatrix(.rows - 1, 14) = rTabela("varSituacaoPGTO")
            
            .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("varDiasJuros"))
            .TextMatrix(.rows - 1, 12) = Format(rTabela("varValorJuros"), ocMONEY)
            
            If .TextMatrix(i, 14) <> "PAGO" Then
               .TextMatrix(.rows - 1, 13) = Format(CDbl(.TextMatrix(.rows - 1, 10)) + CDbl(.TextMatrix(.rows - 1, 12)), ocMONEY)
            Else
               .TextMatrix(.rows - 1, 13) = Format(rTabela("valor_final"), ocMONEY)
            End If
            
            .TextMatrix(.rows - 1, 15) = Format(rTabela("var_DataPgto"), "dd/mm/yy")
            
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
         .Col = 13
         .CellFontBold = True
      Next
      
      'Mudar as cores dependendo da situação
      For i = 1 To .rows - 1
         For j = 0 To .Cols - 1
            .Col = j
            .Row = i
            
            If .TextMatrix(i, 14) <> "PAGO" Then
               If CDate(.TextMatrix(i, 7)) < Date Then
                  .CellForeColor = vbRed
               ElseIf CDate(.TextMatrix(i, 7)) = Date Then
                  .CellForeColor = &H8000&
               ElseIf CDate(.TextMatrix(i, 7)) > Date Then
                  .CellForeColor = vbBlack
               End If
            ElseIf .TextMatrix(i, 14) = "PAGO" Then
               .CellForeColor = vbBlue
            End If
         Next
      Next
      
      FlexCores &HFFFFFF, &HE0E0E0
      
      .Redraw = True
      .Visible = True
      
      'somar a colunas totais
      lblQuant.Caption = Grid.rows - 1
      lblTotalBruto.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
      lblTotalHaver.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
      lblTotalJuros.Caption = Format(SomaGrid(Grid, 12), ocMONEY)
      lblTotalLiquido.Caption = CCur(lblTotalBruto.Caption) - CCur(lblTotalHaver.Caption)
   End With
   
   picAguarde.Visible = False
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset, ByVal percJuros As Currency)
Dim i As Integer, j As Integer

picAguarde.Visible = True
DoEvents

With Grid
   .Visible = False
   
   .Clear
   .Cols = 14
   .rows = 2
   
   .ColWidth(0) = 650
   .ColWidth(1) = 750
   .ColWidth(2) = 4700
   .ColWidth(3) = 750
   .ColWidth(4) = 1250
   .ColWidth(5) = 750
   .ColWidth(6) = 670
   .ColWidth(7) = 450
   .ColWidth(8) = 700
   .ColWidth(9) = 700
   .ColWidth(10) = 670
   .ColWidth(11) = 700
   .ColWidth(12) = 720
   .ColWidth(13) = 750
   
   .TextMatrix(0, 0) = "Cód."
   .TextMatrix(0, 1) = "Origem"
   .TextMatrix(0, 2) = "Cliente"
   .TextMatrix(0, 3) = "Tipo"
   .TextMatrix(0, 4) = "Forma"
   .TextMatrix(0, 5) = "Venc"
   .TextMatrix(0, 6) = "Valor"
   .TextMatrix(0, 7) = "Dias"
   .TextMatrix(0, 8) = "Juros"
   .TextMatrix(0, 9) = "SubTotal"
   .TextMatrix(0, 10) = "Haver"
   .TextMatrix(0, 11) = "Deve"
   .TextMatrix(0, 12) = "Status"
   .TextMatrix(0, 13) = "Pgto"
   
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
   
   .Redraw = False
   i = 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
      
        .TextMatrix(.rows - 1, 0) = ValidateNull(rTabela("var_codigo"))
        .TextMatrix(.rows - 1, 1) = rTabela("var_tipo")
        .TextMatrix(.rows - 1, 2) = rTabela("nome") & IIf(Trim(ValidateNull(rTabela("celular"))) = "", "", "     (" & Right$(ValidateNull(rTabela("celular")), 9) & ")")
        .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("tipo_pagamento"))
        .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_FormaPgto"))
        .TextMatrix(.rows - 1, 5) = Format(rTabela("data"), "dd/mm/yy")
        
        .TextMatrix(.rows - 1, 6) = Format(rTabela("varValorParc"), ocMONEY)
        .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("varDiasJuros"))
        .TextMatrix(.rows - 1, 8) = Format(rTabela("varValorJuros"), ocMONEY)
        .TextMatrix(.rows - 1, 9) = Format(rTabela("varParcComJuros"), ocMONEY)
        
        .TextMatrix(.rows - 1, 10) = Format(rTabela("var_haver"), ocMONEY)
        .TextMatrix(.rows - 1, 12) = rTabela("varSituacaoPGTO")
        .TextMatrix(.rows - 1, 13) = Format(rTabela("var_DataPgto"), "dd/mm/yy")
        
        'If .TextMatrix(i, 11) <> "PAGO" Then
        '   .TextMatrix(.rows - 1, 11) = Format(CDbl(.TextMatrix(.rows - 1, 9) - CDbl(.TextMatrix(.rows - 1, 8))), ocMONEY)
        'Else
           .TextMatrix(.rows - 1, 11) = Format(rTabela("varTotalLiquido"), ocMONEY)
        'End If
        
        rTabela.MoveNext
        .rows = .rows + 1
        i = i + 1
      Loop
   End If
   
   .rows = .rows - 1

   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 3
      .CellFontBold = True
   Next
   
   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .rows - 1
      .Row = i
      .Col = 13
      .CellFontBold = True
   Next
   
   'Mudar as cores dependendo da situação
   For i = 1 To .rows - 1
      For j = 0 To .Cols - 1
         .Col = j
         .Row = i
         
         If .TextMatrix(i, 11) <> "PAGO" Then
            If CDate(.TextMatrix(i, 5)) < Date Then
               .CellForeColor = vbRed
            ElseIf CDate(.TextMatrix(i, 5)) = Date Then
               .CellForeColor = &H8000&
            ElseIf CDate(.TextMatrix(i, 5)) > Date Then
               .CellForeColor = vbBlack
            End If
         ElseIf .TextMatrix(i, 12) = "PAGO" Then
            .CellForeColor = vbBlue
         End If
      Next
   Next
   
   FlexCores &HFFFFFF, &HE0E0E0
   
   .Redraw = True
   .Visible = True
   
   'somar a colunas totais
   lblQuant.Caption = Grid.rows - 1
   lblTotalParcelas.Caption = Format(SomaGrid(Grid, 6), ocMONEY)
   lblTotalBruto.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
   lblTotalHaver.Caption = Format(SomaGrid(Grid, 10), ocMONEY)
   lblTotalJuros.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
   lblTotalLiquido.Caption = CCur(lblTotalBruto.Caption) - CCur(lblTotalHaver.Caption)
End With

picAguarde.Visible = False
End Sub


Function EImpar(ByVal iNum As Long) As Boolean
   EImpar = (iNum Mod 2)
End Function

Sub FlexCores(lCorPar As Long, lCorImpar As Long)
   'ZEBRAR O FLEXGRID
   Dim iLinha As Integer
   Dim lCor As OLE_COLOR
   
   Grid.FillStyle = flexFillRepeat
   
   For iLinha = 1 To Grid.rows - 1
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

Private Sub cmdImprime_Click()
If cboTipoImpressao.Text = "" Then Exit Sub

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = vImpressoraNormal Then
      Set Printer = Prt
      Exit For
   End If
Next

Me.Hide

If cboTipoImpressao.Text = "RELATÓRIO NORMAL" Then
        cmdExibir_Click
        Set r = dbData.OpenRecordset(printSQL)
        Set REL_Parcelas_Cons.Relatorio.Recordset = r
        REL_Parcelas_Cons.rfQuant.Caption = lblQuant.Caption
        REL_Parcelas_Cons.rfParcelas.Caption = Format(lblTotalParcelas.Caption, "##,##0.00")
        REL_Parcelas_Cons.rfSubTotal.Caption = Format(lblTotalBruto.Caption, "##,##0.00")
        REL_Parcelas_Cons.rfJuros.Caption = Format(lblTotalJuros.Caption, "##,##0.00")
        REL_Parcelas_Cons.rfHaveres.Caption = Format(lblTotalHaver.Caption, "##,##0.00")
        REL_Parcelas_Cons.rfTotal.Caption = Format(lblTotalLiquido.Caption, "##,##0.00")
        
        REL_Parcelas_Cons.rfForma.Caption = cboForma.Text
        REL_Parcelas_Cons.rfTipo.Caption = cboTipo.Text
        
        If cboCriterios.Text = "CLIENTE" Then
           REL_Parcelas_Cons.rfCons1.Caption = "Cliente = " & cboNome.Text & ""
        ElseIf cboCriterios.Text = "VENDEDOR" Then
           REL_Parcelas_Cons.rfCons1.Caption = "Vendedor = " & cboNome.Text & ""
        ElseIf cboCriterios.Text = "DATA" Then
           REL_Parcelas_Cons.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
        ElseIf cboCriterios.Text = "MENSAL" Then
           REL_Parcelas_Cons.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
        ElseIf cboCriterios.Text = "TODOS" Then
           REL_Parcelas_Cons.rfCons1.Caption = "TODOS"
        End If
        
        If cboSituacao.Text = "VENCIDAS" Or cboSituacao.Text = "À PAGAR" Then
            REL_Parcelas_Cons.lblPGTO.Visible = False
        Else
            REL_Parcelas_Cons.lblPGTO.Visible = True
        End If
        
        REL_Parcelas_Cons.Relatorio.NomeImpressora = vImpressoraNormal
        REL_Parcelas_Cons.Relatorio.Ativar
        Unload REL_Parcelas_Cons
        Me.Show 1
ElseIf cboTipoImpressao.Text = "RELATÓRIO AGRUPADO" Then

   'If cboCriterios.Text = "TODOS" Then
   'ElseIf cboCriterios.Text = "DATA" Then
   'ElseIf cboCriterios.Text = "MENSAL" Then
   'ElseIf cboCriterios.Text = "CLIENTE" Then
   'ElseIf cboCriterios.Text = "VENDEDOR" Then
   'End If
   
    sSQL = "SELECT DISTINCT cliente.CODIGO AS varClienteCod, cliente.Nome, CASE WHEN cliente.telefone1 IS NULL THEN cliente.celular WHEN cliente.telefone1 = '' THEN cliente.celular ELSE cliente.telefone1 END AS vContato, " & _
           "(SELECT ISNULL(SUM(parcelas.VALOR), 0) AS Expr1 FROM parcelas INNER JOIN pedidos ON parcelas.COD_PEDIDO = pedidos.COD_PEDIDO WHERE (pedidos.COD_CLIENTE = cliente.CODIGO) AND (parcelas.STATUS = 0) AND (parcelas.DATA < GETDATE())) AS varSomaParcelasVencidas, " & _
           "(SELECT ISNULL(SUM(parcelas_2.VALOR), 0) AS Expr1 FROM parcelas AS parcelas_2 INNER JOIN pedidos AS pedidos_2 ON parcelas_2.COD_PEDIDO = pedidos_2.COD_PEDIDO WHERE (pedidos_2.COD_CLIENTE = cliente.CODIGO) AND (parcelas_2.STATUS = 0) AND (parcelas_2.DATA > GETDATE())) AS varSomaParcelasAVencer " & _
           "FROM parcelas AS parcelas_1 INNER JOIN pedidos AS pedidos_1 ON parcelas_1.COD_PEDIDO = pedidos_1.COD_PEDIDO INNER JOIN cliente ON pedidos_1.COD_CLIENTE = cliente.CODIGO " & _
           "Where (parcelas_1.Status = 0) " & _
           "GROUP BY cliente.CODIGO, cliente.Nome, cliente.TELEFONE1, cliente.CELULAR " & _
           "ORDER BY cliente.Nome"
    Set r = dbData.OpenRecordset(sSQL)
    'Debug.Print sSQL
    printSQL = sSQL
    
    Me.Hide
    
    Set r = dbData.OpenRecordset(printSQL)
    
    Set REL_Parcelas_Agrupado.Relatorio.Recordset = r
    
    'REL_Parcelas_Agrupado.rfQuant.Caption = lblQuant.Caption
    'REL_Parcelas_Agrupado.rfsubtotal.Caption = Format(lblTotalBruto.Caption, "##,##0.00")
    
    'REL_Parcelas_Agrupado.rfForma.Caption = cboForma.Text
    'REL_Parcelas_Agrupado.rfTipo.Caption = cboTipo.Text
    
    REL_Parcelas_Agrupado.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Parcelas_Agrupado.Relatorio.Ativar
    Unload REL_Parcelas_Agrupado
    
   Me.Show 1
ElseIf cboTipoImpressao.Text = "RELATÓRIO UNIFICADO" Then

    If cboCriterios.Text <> "CLIENTE" Then MsgBox "Essa impressão somente é permitida para consulta por cliente!", vbInformation, "Aviso do Sistema": Exit Sub
    If txtCodFunc.Text = "" Then MsgBox "Selecione um cliente!", vbInformation, "Aviso do Sistema": Exit Sub
    cmdExibir_Click
    Me.Hide
    Set r = dbData.OpenRecordset(printSQL)
    Set REL_Parcelas_PorClientes.Relatorio.Recordset = r
    
    REL_Parcelas_PorClientes.rfQuant.Caption = lblQuant.Caption
    REL_Parcelas_PorClientes.rfParcelas.Caption = Format(lblTotalParcelas.Caption, "##,##0.00")
    REL_Parcelas_PorClientes.rfSubTotal.Caption = Format(lblTotalBruto.Caption, "##,##0.00")
    REL_Parcelas_PorClientes.rfJuros.Caption = Format(lblTotalJuros.Caption, "##,##0.00")
    REL_Parcelas_PorClientes.rfHaveres.Caption = Format(lblTotalHaver.Caption, "##,##0.00")
    REL_Parcelas_PorClientes.rfTotal.Caption = Format(lblTotalLiquido.Caption, "##,##0.00")
    
    REL_Parcelas_PorClientes.rfForma.Caption = cboSituacao.Text
    REL_Parcelas_PorClientes.rfCons1.Caption = "Cliente = " & cboNome.Text & ""
    
    REL_Parcelas_PorClientes.rfNome.Caption = vNome
    REL_Parcelas_PorClientes.rfCelular.Caption = vCelular
    REL_Parcelas_PorClientes.rfEndereco.Caption = vEndereco
    REL_Parcelas_PorClientes.rfNum.Caption = vNum
    REL_Parcelas_PorClientes.rfBairro.Caption = vBairro
    REL_Parcelas_PorClientes.rfCidade.Caption = vCidade
    REL_Parcelas_PorClientes.rfUF.Caption = vUF
    REL_Parcelas_PorClientes.rfReferencia.Caption = vReferencia
    
    If cboSituacao.Text = "VENCIDAS" Or cboSituacao.Text = "À PAGAR" Then
        REL_Parcelas_Cons.lblPGTO.Visible = False
    Else
        REL_Parcelas_Cons.lblPGTO.Visible = True
    End If
    
    REL_Parcelas_PorClientes.Relatorio.NomeImpressora = vImpressoraNormal
    REL_Parcelas_PorClientes.Relatorio.Ativar
    Unload REL_Parcelas_PorClientes
    Me.Show 1
Else
    Exit Sub
End If
End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset
Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Parcelas_Cons.Relatorio.Recordset = r

REL_Parcelas_Cons.rfQuant.Caption = lblQuant.Caption
REL_Parcelas_Cons.rfParcelas.Caption = Format(lblTotalParcelas.Caption, "##,##0.00")
REL_Parcelas_Cons.rfSubTotal.Caption = Format(lblTotalBruto.Caption, "##,##0.00")
REL_Parcelas_Cons.rfJuros.Caption = Format(lblTotalJuros.Caption, "##,##0.00")
REL_Parcelas_Cons.rfHaveres.Caption = Format(lblTotalHaver.Caption, "##,##0.00")
REL_Parcelas_Cons.rfTotal.Caption = Format(lblTotalLiquido.Caption, "##,##0.00")

REL_Parcelas_Cons.rfForma.Caption = cboForma.Text
REL_Parcelas_Cons.rfTipo.Caption = cboTipo.Text

If cboCriterios.Text = "CLIENTE" Then
   REL_Parcelas_Cons.rfCons1.Caption = "Cliente = " & cboNome.Text & ""
ElseIf cboCriterios.Text = "VENDEDOR" Then
   REL_Parcelas_Cons.rfCons1.Caption = "Vendedor = " & cboNome.Text & ""
ElseIf cboCriterios.Text = "DATA" Then
   REL_Parcelas_Cons.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
ElseIf cboCriterios.Text = "MENSAL" Then
   REL_Parcelas_Cons.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
ElseIf cboCriterios.Text = "TODOS" Then
   REL_Parcelas_Cons.rfCons1.Caption = "TODOS"
End If

REL_Parcelas_Cons.Relatorio.NomeImpressora = vImpressoraNormal
REL_Parcelas_Cons.Relatorio.Ativar
Unload REL_Parcelas_Cons

Me.Show 1
End Sub

Private Sub Form_Load()
Preencher_Parcelas
Preencher_TipoCartao
Preencher_SubCriterios
Preencher_Setor
Preencher_Forma
Preencher_Criterio
Preencher_Ordem
Preencher_Situacao
Preencher_Tipo
'Preencher_Datas

cboSetor.ListIndex = 1
cboSituacao.ListIndex = 3
cboCriterios.ListIndex = 0
cboOrdem.ListIndex = 3
cboForma.ListIndex = 2
cboTipo.ListIndex = 0
'cboDatas.ListIndex = 1
'cboJuros.ListIndex = 0
'cboSubCriterios.ListIndex = 0
'cboTipoCartao.ListIndex = 0
'cboParcelas.ListIndex = 1

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
vImpressoraNormal = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Limpar_Grid
Set moCombo = New cComboHelper
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub mskFim_KeyPress(KeyAscii As Integer)
   mskFim.Mask = "##/##/##"
End Sub

Private Sub mskFim_LostFocus()
   If mskFim.Text = "" Or mskFim.Text = "__/__/__" Then
      mskFim.Mask = ""
      mskFim.Text = ""
      Exit Sub
   Else
      If IsDate(mskFim.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskFim.SetFocus
         SelectControl mskFim
      End If
   End If
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub

Private Sub mskInicio_LostFocus()
   If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
      mskInicio.Mask = ""
      mskInicio.Text = ""
      Exit Sub
   Else
      If IsDate(mskInicio.Text) Then
         If mskFim.Enabled = True Then mskFim.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskInicio.SetFocus
         SelectControl mskInicio
      End If
   End If
End Sub


Private Sub OptDataIntervalo_Click()
If optDataUnico.Value = True Then
   lblInicio.Enabled = True
   mskInicio.Enabled = True
   lblFim.Enabled = False
   mskFim.Enabled = False
   cmdCal2.Enabled = False
   mskInicio.SetFocus
ElseIf OptDataIntervalo.Value = True Then
   lblInicio.Enabled = True
   mskInicio.Enabled = True
   lblFim.Enabled = True
   mskFim.Enabled = True
   cmdCal2.Enabled = True
   mskInicio.SetFocus
End If
End Sub

Private Sub optDataUnico_Click()
   If optDataUnico.Value = True Then
      lblInicio.Enabled = True
      mskInicio.Enabled = True
      lblFim.Enabled = False
      mskFim.Enabled = False
      cmdCal2.Enabled = False
      mskInicio.SetFocus
   ElseIf OptDataIntervalo.Value = True Then
      lblInicio.Enabled = True
      mskInicio.Enabled = True
      lblFim.Enabled = True
      mskFim.Enabled = True
      cmdCal2.Enabled = True
      mskInicio.SetFocus
   End If
End Sub

Private Sub txtCodFunc_Change()
If txtCodFunc.Text = "" Then Exit Sub
If cboCriterios.Text = "CLIENTE" Then
    sSQL = "SELECT * FROM cliente WHERE (codigo= " & txtCodFunc.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.BOF Then
        
        vNome = r("nome")
        vCelular = r("celular")
        vEndereco = r("endereco")
        vNum = r("numero")
        vBairro = r("bairro")
        vCidade = r("cidade")
        vUF = r("estado")
        vReferencia = r("Ponto_de_referencia")
    If r.State <> 0 Then r.Close
    Set r = Nothing
    End If
End If
End Sub


