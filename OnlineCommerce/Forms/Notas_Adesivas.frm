VERSION 5.00
Begin VB.Form Notas_Adesivas 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RESUMO FISCAL"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   60
      TabIndex        =   37
      Top             =   5940
      Width           =   4455
      Begin VB.OptionButton optMensal 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mensal"
         Height          =   195
         Left            =   1800
         TabIndex        =   40
         Top             =   360
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optAnual 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Anual"
         Height          =   195
         Left            =   2700
         TabIndex        =   39
         Top             =   360
         Width           =   915
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de consulta:"
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
         Left            =   120
         TabIndex        =   46
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mês de Referência:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   360
         TabIndex        =   45
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblMesRef 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1920
         TabIndex        =   44
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lblAvancar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2040
         TabIndex        =   43
         Top             =   600
         Width           =   240
      End
      Begin VB.Label lblVoltar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   240
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
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   345
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2640
      Top             =   3600
   End
   Begin VB.Image imgExcl4 
      Height          =   240
      Left            =   4140
      Picture         =   "Notas_Adesivas.frx":0000
      Top             =   5460
      Width           =   270
   End
   Begin VB.Image imgExcl3 
      Height          =   240
      Left            =   4140
      Picture         =   "Notas_Adesivas.frx":0674
      Top             =   3660
      Width           =   270
   End
   Begin VB.Image imgExcl2 
      Height          =   240
      Left            =   4140
      Picture         =   "Notas_Adesivas.frx":0CE8
      Top             =   2580
      Width           =   270
   End
   Begin VB.Image imgExcl1 
      Height          =   240
      Left            =   4140
      Picture         =   "Notas_Adesivas.frx":135C
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Para alcançar a meta 2 falta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   36
      Top             =   5640
      Width           =   2445
   End
   Begin VB.Label lblVlrRestante2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3705
      TabIndex        =   35
      Top             =   5640
      Width           =   375
   End
   Begin VB.Line Line4 
      X1              =   60
      X2              =   4140
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emissão Fiscal até o momento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   34
      Top             =   5160
      Width           =   2640
   End
   Begin VB.Label lblVlrSaida2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   3705
      TabIndex        =   33
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "META 3: Maior valor (meta 1 e 2):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   32
      Top             =   4800
      Width           =   2925
   End
   Begin VB.Label lblMeta1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   3700
      TabIndex        =   31
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "META 2: Recebimentos (vendas):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   30
      Top             =   4560
      Width           =   2910
   End
   Begin VB.Label lblEmitido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   29
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblMeta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3700
      TabIndex        =   28
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "META FISCAL:"
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
      Left            =   60
      TabIndex        =   27
      Top             =   4020
      Width           =   1365
   End
   Begin VB.Line Line3 
      X1              =   3480
      X2              =   4080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   4080
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EMISSÃO FISCAL (vendas):"
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
      Left            =   60
      TabIndex        =   26
      Top             =   2940
      Width           =   2580
   End
   Begin VB.Label lblTotalRecebido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   25
      Top             =   1560
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   24
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label lblRecDepos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   23
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label31 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cartão Crédito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   22
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label lblRecTransf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   21
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblRecCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   20
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Transferência:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RECEBIMENTOS:"
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
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   1590
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   17
      Top             =   360
      Width           =   285
   End
   Begin VB.Label lblRecDebito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   16
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblRecPix 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3700
      TabIndex        =   15
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cartão Débito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   14
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label lblVlrRestante 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3705
      TabIndex        =   13
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Para alcançar a meta 1 falta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   12
      Top             =   5400
      Width           =   2445
   End
   Begin VB.Label lblVlrSaida 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   3700
      TabIndex        =   11
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblVlrEntrada 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   3700
      TabIndex        =   10
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "META 1: Compras + 25%:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblNFCe 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   3700
      TabIndex        =   8
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lblNFe 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   3700
      TabIndex        =   7
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NFCe Emitidas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   3480
      Width           =   1380
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada Manual:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label lblEntradaXML 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   3700
      TabIndex        =   4
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblEntradaManual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   3700
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada XML:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NFe Emitidas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      TabIndex        =   1
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTAS FISCAIS (compras):"
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
      Left            =   60
      TabIndex        =   0
      Top             =   1860
      Width           =   2565
   End
End
Attribute VB_Name = "Notas_Adesivas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vMesInt As String
Dim vAno As String
Dim vNumeroMes As Integer
Dim vTotalXML As Currency
Dim vTotalManual As Currency
Dim vTotalPix As Currency
Dim vTotalCartaoC As Currency
Dim vTotalCartaoD As Currency
Dim vTotalTransf As Currency
Dim vTotalDeposito As Currency
Dim vDataHoje As Date

Private Sub Contas_Produtos()
'sSQL = "SELECT codigo FROM produtos WHERE ativo = 0 and quant_min >= quant_estoque;"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'    lblEstoqueMin.Caption = Format(0, "000")
'Else
'    lblEstoqueMin.Caption = Format(r.RecordCount, "000")
'End If

'sSQL = "SELECT codigo FROM produtos_comprar;"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'    lblLista.Caption = Format(0, "000")
'Else
'    lblLista.Caption = Format(r.RecordCount, "000")
'End If

'If r.State <> 0 Then r.Close
'Set r = Nothing
End Sub

Private Sub Contas_AReceber()
'sSQL = "SELECT codigo FROM parcelas WHERE status = 0 and (data <= CONVERT(DATETIME, '" & Format(Date, ocDATA) & "'));"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'    lblClientes.Caption = Format(0, "000")
'Else
'    lblClientes.Caption = Format(r.RecordCount, "000")
'End If

'sSQL = "SELECT codigo FROM parcelas WHERE status = 0 and (data = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "'));"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'    lblClientesHJ.Caption = Format(0, "000")
'Else
'    lblClientesHJ.Caption = Format(r.RecordCount, "000")
'End If

'If r.State <> 0 Then r.Close
'Set r = Nothing
End Sub

Private Sub Contas_Apagar()
'sSQL = "SELECT codigo FROM a_pagar WHERE status = 'À PAGAR';"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'    lblContas.Caption = Format(0, "000")
'Else
'    lblContas.Caption = Format(r.RecordCount, "000")
'End If

'sSQL = "SELECT codigo FROM a_pagar WHERE status = 'À PAGAR' and (vencimento = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "'));"
'Set r = dbData.OpenRecordset(sSQL)

'If r.BOF Or r.RecordCount = 0 Then
'    lblContasHJ.Caption = Format(0, "000")
'Else
'    lblContasHJ.Caption = Format(r.RecordCount, "000")
'End If

'If r.State <> 0 Then r.Close
'Set r = Nothing
End Sub
Private Sub Mostrar_Entradas()
'valor total das xml
If optMensal.Value = True Then
    sSQL = "SELECT SUM(ValorNota) AS vTotalXML FROM EntradaEstoque WHERE (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ");"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(ValorNota) AS vTotalXML FROM EntradaEstoque WHERE (YEAR(DataEmissao) = " & cboAno & ");"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblEntradaXML.Caption = Format(ValidateNull(r("vTotalXML")), ocMONEY)
    vTotalXML = Format(ValidateNull(r("vTotalXML")), ocMONEY)
Else
    lblEntradaXML.Caption = Format(0, ocMONEY)
    vTotalXML = Format(0, ocMONEY)
End If

'valor total entradas
'sSQL = "SELECT SUM(ValorNota) AS vTotalXML FROM EntradaEstoque WHERE (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ");"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.BOF Then
'    lblEntradaManual.Caption = Format(ValidateNull(r("vTotalXML")), ocMONEY)
'Else
    lblEntradaManual.Caption = Format(0, ocMONEY)
    vTotalManual = Format(0, ocMONEY)
'End If
End Sub

Private Sub Mostrar_RecDeposito()
'PIX
'parcelas
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (MONTH(PAGAMENTO) = " & vNumeroMes & ") AND (YEAR(PAGAMENTO) = " & vAno & ") and STATUS = 1 AND FORMA_PGTO = 'DEPOSITO' ;"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (YEAR(PAGAMENTO) = " & cboAno & ") and STATUS = 1 AND FORMA_PGTO = 'DEPOSITO' ;"
End If
Set r = dbData.OpenRecordset(sSQL)

Dim vParcPix As Currency
Dim vHaverPix As Currency

If Not r.BOF Then
    vParcPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vParcPix = FormatNumber(0, 2)
End If

'haver
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (MONTH(HAVER) = " & vNumeroMes & ") AND (YEAR(HAVER) = " & vAno & ") and FORMA_PGTO = 'DEPOSITO' ;"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (YEAR(HAVER) = " & cboAno & ") and FORMA_PGTO = 'DEPOSITO' ;"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    vHaverPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vHaverPix = FormatNumber(0, 2)
End If
lblRecDepos.Caption = FormatNumber(vParcPix + vHaverPix, 2)
vTotalDeposito = FormatNumber(lblRecDepos.Caption, 2)
End Sub

Private Sub Mostrar_RecTransf()
'PIX
'parcelas
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (MONTH(PAGAMENTO) = " & vNumeroMes & ") AND (YEAR(PAGAMENTO) = " & vAno & ") and STATUS = 1 AND FORMA_PGTO = 'TRANSFERENCIA' ;"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (YEAR(PAGAMENTO) = " & cboAno & ") and STATUS = 1 AND FORMA_PGTO = 'TRANSFERENCIA' ;"
End If
Set r = dbData.OpenRecordset(sSQL)

Dim vParcPix As Currency
Dim vHaverPix As Currency

If Not r.BOF Then
    vParcPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vParcPix = FormatNumber(0, 2)
End If

'haver
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (MONTH(HAVER) = " & vNumeroMes & ") AND (YEAR(HAVER) = " & vAno & ") and FORMA_PGTO = 'TRANSFERENCIA' ;"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (YEAR(HAVER) = " & cboAno & ") and FORMA_PGTO = 'TRANSFERENCIA' ;"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    vHaverPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vHaverPix = FormatNumber(0, 2)
End If
lblRecTransf.Caption = FormatNumber(vParcPix + vHaverPix, 2)
vTotalTransf = FormatNumber(lblRecTransf.Caption, 2)
End Sub

Private Sub Mostrar_RecCartaoD()
'PIX
'parcelas
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (MONTH(PAGAMENTO) = " & vNumeroMes & ") AND (YEAR(PAGAMENTO) = " & vAno & ") and STATUS = 1 AND FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'D';"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (YEAR(PAGAMENTO) = " & cboAno & ") and STATUS = 1 AND FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'D';"
End If
Set r = dbData.OpenRecordset(sSQL)

Dim vParcPix As Currency
Dim vHaverPix As Currency

If Not r.BOF Then
    vParcPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vParcPix = FormatNumber(0, 2)
End If

'haver
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (MONTH(HAVER) = " & vNumeroMes & ") AND (YEAR(HAVER) = " & vAno & ") and FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'D';"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (YEAR(HAVER) = " & cboAno & ") and FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'D';"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    vHaverPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vHaverPix = FormatNumber(0, 2)
End If
lblRecDebito.Caption = FormatNumber(vParcPix + vHaverPix, 2)
vTotalCartaoD = FormatNumber(lblRecDebito.Caption, 2)
End Sub

Private Sub Mostrar_RecCartaoC()
'PIX
'parcelas
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (MONTH(PAGAMENTO) = " & vNumeroMes & ") AND (YEAR(PAGAMENTO) = " & vAno & ") and STATUS = 1 AND FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'C';"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalCartaoC FROM parcelas WHERE (YEAR(PAGAMENTO) = " & cboAno & ") and STATUS = 1 AND FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'C';"
End If
Set r = dbData.OpenRecordset(sSQL)

Dim vParcPix As Currency
Dim vHaverPix As Currency

If Not r.BOF Then
    vParcPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vParcPix = FormatNumber(0, 2)
End If

'haver
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (MONTH(HAVER) = " & vNumeroMes & ") AND (YEAR(HAVER) = " & vAno & ") and FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'C';"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_HAVER) AS varTotalCartaoC FROM parcelas_haver WHERE (YEAR(HAVER) = " & cboAno & ") and FORMA_PGTO = 'CARTAO' and TIPO_CARTAO = 'C';"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    vHaverPix = FormatNumber(ValidateNull(r("varTotalCartaoC")), 2)
Else
    vHaverPix = FormatNumber(0, 2)
End If
lblRecCredito.Caption = FormatNumber(vParcPix + vHaverPix, 2)
vTotalCartaoC = FormatNumber(lblRecCredito.Caption, 2)
End Sub

Private Sub Mostrar_RecPIX()
'PIX
'parcelas
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalPIX FROM parcelas WHERE (MONTH(PAGAMENTO) = " & vNumeroMes & ") AND (YEAR(PAGAMENTO) = " & vAno & ") and STATUS = 1 AND FORMA_PGTO = 'PIX';"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_FINAL) AS varTotalPIX FROM parcelas WHERE (YEAR(PAGAMENTO) = " & cboAno & ") and STATUS = 1 AND FORMA_PGTO = 'PIX';"
End If
Set r = dbData.OpenRecordset(sSQL)

Dim vParcPix As Currency
Dim vHaverPix As Currency

If Not r.BOF Then
    vParcPix = FormatNumber(ValidateNull(r("varTotalPIX")), 2)
Else
    vParcPix = FormatNumber(0, 2)
End If

'haver
If optMensal.Value = True Then
    sSQL = "SELECT SUM(VALOR_HAVER) AS varHaverPIX FROM parcelas_haver WHERE (MONTH(HAVER) = " & vNumeroMes & ") AND (YEAR(HAVER) = " & vAno & ") and FORMA_PGTO = 'PIX';"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(VALOR_HAVER) AS varHaverPIX FROM parcelas_haver WHERE (YEAR(HAVER) = " & cboAno & ") and FORMA_PGTO = 'PIX';"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    vHaverPix = FormatNumber(ValidateNull(r("varHaverPIX")), 2)
Else
    vHaverPix = FormatNumber(0, 2)
End If
lblRecPix.Caption = FormatNumber(vParcPix + vHaverPix, 2)
vTotalPix = FormatNumber(lblRecPix.Caption, 2)
End Sub

Private Sub cboAno_Click()
cboAno_LostFocus
End Sub

Private Sub cboAno_GotFocus()
'moCombo.AttachTo cboAno
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
' O último ano será o anterior ao atual
LastYear = iAno
' O primeiro ano será 5 anos atrás em relação ao LastYear
FirstYear = iAno - 5

For i = FirstYear To LastYear
   cboAno.AddItem CStr(i)
Next

' Se o moCombo for um componente de subclasse ou skin, mantenha a linha abaixo
' moCombo.AttachTo cboAno

' Opcional: Deixar o ano mais recente (ano passado) já selecionado
If cboAno.ListCount > 0 Then
    cboAno.ListIndex = cboAno.ListCount - 1
End If

End Sub


Private Sub cboAno_LostFocus()
If Not IsNumeric(cboAno.Text) Or Len(cboAno.Text) <> 4 Then
    MsgBox "Por favor, digite um ano válido com 4 dígitos (ex: 2026).", vbExclamation, "Erro de Data"
    cboAno.SetFocus
Else
    Dim ano As Integer
    ano = CInt(cboAno.Text)
    
    If ano < 1900 Or ano > 2100 Then
        MsgBox "O ano deve estar entre 1900 e 2100.", vbExclamation, "Ano Inválido"
        cboAno.SetFocus
    Else
        Form_Load
    End If
End If
End Sub


Private Sub cboAno_Validate(Cancel As Boolean)
cboAno_LostFocus
End Sub

Private Sub Form_Activate()
cboAno.Text = Year(Date)
End Sub

Private Sub Form_Load()
lblVoltar.Visible = True
Label14.Visible = True
lblAvancar.Visible = True
lblMesRef.Visible = True
If optAnual.Value = False Then
    lblAno.Visible = False
    cboAno.Visible = False
End If
'definição de mês
vDataHoje = Date

vMesInt = Format(Date, "mmmm")
vAno = Year(Date)
lblMesRef.Caption = vMesInt & "/" & vAno

If vMesInt = "janeiro" Then
    vNumeroMes = 1
ElseIf vMesInt = "fevereiro" Then
    vNumeroMes = 2
ElseIf vMesInt = "março" Then
    vNumeroMes = 3
ElseIf vMesInt = "abril" Then
    vNumeroMes = 4
ElseIf vMesInt = "maio" Then
    vNumeroMes = 5
ElseIf vMesInt = "junho" Then
    vNumeroMes = 6
ElseIf vMesInt = "julho" Then
    vNumeroMes = 7
ElseIf vMesInt = "agosto" Then
    vNumeroMes = 8
ElseIf vMesInt = "setembro" Then
    vNumeroMes = 9
ElseIf vMesInt = "outubro" Then
    vNumeroMes = 10
ElseIf vMesInt = "novembro" Then
    vNumeroMes = 11
ElseIf vMesInt = "dezembro" Then
    vNumeroMes = 12
End If

Mostrar_RecPIX
Mostrar_RecCartaoC
Mostrar_RecCartaoD
Mostrar_RecDeposito
Mostrar_RecTransf

Dim vTotalRecebido As Currency
vTotalRecebido = vTotalPix + vTotalCartaoC + vTotalCartaoD + vTotalTransf + vTotalDeposito
lblTotalRecebido.Caption = FormatNumber(vTotalRecebido, 2)

Mostrar_Entradas

'valor NFCE
Dim vTotalNFCE As Currency
If optMensal.Value = True Then
    sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (YEAR(DataEmissao) = " & cboAno & ")"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblNFCe.Caption = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
    vTotalNFCE = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
Else
    lblNFCe.Caption = FormatNumber(0, 2)
    vTotalNFCE = FormatNumber(0, 2)
End If

'valor NFE
Dim vTotalNFe As Currency
If optMensal.Value = True Then
    sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (YEAR(DataEmissao) = " & cboAno & ")"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblNFe.Caption = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
    vTotalNFe = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
Else
    lblNFe.Caption = FormatNumber(0, 2)
    vTotalNFe = FormatNumber(0, 2)
End If

'somar totais
Dim vTotalEntradas As Currency
Dim vTotalSaidas As Currency
Dim vRestante As Currency

vTotalEntradas = vTotalXML + vTotalManual
vTotalSaidas = vTotalNFCE + vTotalNFe

lblVlrEntrada.Caption = FormatNumber(vTotalEntradas, 2)
lblVlrSaida.Caption = FormatNumber(vTotalSaidas, 2)
lblVlrSaida2.Caption = FormatNumber(vTotalSaidas, 2)

vTotalEntradas = vTotalEntradas + (vTotalEntradas * 0.25)
lblMeta1.Caption = FormatNumber(vTotalEntradas, 2)  'META 1

If vTotalRecebido > vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalRecebido, 2)
ElseIf vTotalRecebido < vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
ElseIf vTotalRecebido = vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
End If
Dim vMeta As Currency
vMeta = FormatNumber(lblMeta.Caption, 2)            'META 3

lblEmitido.Caption = FormatNumber(vTotalRecebido, 2)  'META 2

vRestante = vMeta - vTotalSaidas

If vRestante > 0 Then
    lblVlrRestante.Caption = FormatNumber(vTotalEntradas - vTotalSaidas, 2)
    If vTotalRecebido > 0 Then
        lblVlrRestante2.Caption = FormatNumber(vTotalRecebido - vTotalSaidas, 2)
    Else
        lblVlrRestante2.Caption = FormatNumber(0, 2)
    End If
    'lblVlrRestante.Caption = FormatNumber(vRestante, 2)
Else
    lblVlrRestante.Caption = FormatNumber(0, 2)
    lblVlrRestante2.Caption = FormatNumber(0, 2)
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub lblAvancar_Click()
'definição de mês
vDataHoje = Format(DateAdd("m", Val(1), vDataHoje), "dd/mm/yy")

vMesInt = Format(vDataHoje, "mmmm")
vAno = Year(vDataHoje)
lblMesRef.Caption = vMesInt & "/" & vAno

If vMesInt = "janeiro" Then
    vNumeroMes = 1
ElseIf vMesInt = "fevereiro" Then
    vNumeroMes = 2
ElseIf vMesInt = "março" Then
    vNumeroMes = 3
ElseIf vMesInt = "abril" Then
    vNumeroMes = 4
ElseIf vMesInt = "maio" Then
    vNumeroMes = 5
ElseIf vMesInt = "junho" Then
    vNumeroMes = 6
ElseIf vMesInt = "julho" Then
    vNumeroMes = 7
ElseIf vMesInt = "agosto" Then
    vNumeroMes = 8
ElseIf vMesInt = "setembro" Then
    vNumeroMes = 9
ElseIf vMesInt = "outubro" Then
    vNumeroMes = 10
ElseIf vMesInt = "novembro" Then
    vNumeroMes = 11
ElseIf vMesInt = "dezembro" Then
    vNumeroMes = 12
End If

Mostrar_RecPIX
Mostrar_RecCartaoC
Mostrar_RecCartaoD
Mostrar_RecDeposito
Mostrar_RecTransf

Dim vTotalRecebido As Currency
vTotalRecebido = vTotalPix + vTotalCartaoC + vTotalCartaoD + vTotalTransf + vTotalDeposito
lblTotalRecebido.Caption = FormatNumber(vTotalRecebido, 2)

Mostrar_Entradas

'valor NFCE
Dim vTotalNFCE As Currency
If optMensal.Value = True Then
    sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (YEAR(DataEmissao) = " & cboAno & ")"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblNFCe.Caption = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
    vTotalNFCE = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
Else
    lblNFCe.Caption = FormatNumber(0, 2)
    vTotalNFCE = FormatNumber(0, 2)
End If

'valor NFE
Dim vTotalNFe As Currency
If optMensal.Value = True Then
    sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (YEAR(DataEmissao) = " & cboAno & ")"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblNFe.Caption = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
    vTotalNFe = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
Else
    lblNFe.Caption = FormatNumber(0, 2)
    vTotalNFe = FormatNumber(0, 2)
End If

'somar totais
Dim vTotalEntradas As Currency
Dim vTotalSaidas As Currency
Dim vRestante As Currency

vTotalEntradas = vTotalXML + vTotalManual
vTotalSaidas = vTotalNFCE + vTotalNFe

lblVlrEntrada.Caption = FormatNumber(vTotalEntradas, 2)
lblVlrSaida.Caption = FormatNumber(vTotalSaidas, 2)
lblVlrSaida2.Caption = FormatNumber(vTotalSaidas, 2)

vTotalEntradas = vTotalEntradas + (vTotalEntradas * 0.25)
lblMeta1.Caption = FormatNumber(vTotalEntradas, 2)  'META 1

If vTotalRecebido > vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalRecebido, 2)
ElseIf vTotalRecebido < vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
ElseIf vTotalRecebido = vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
End If
Dim vMeta As Currency
vMeta = FormatNumber(lblMeta.Caption, 2)            'META 3

lblEmitido.Caption = FormatNumber(vTotalRecebido, 2)  'META 2

vRestante = vMeta - vTotalSaidas

If vRestante > 0 Then
    lblVlrRestante.Caption = FormatNumber(vTotalEntradas - vTotalSaidas, 2)
    If vTotalRecebido > 0 Then
        lblVlrRestante2.Caption = FormatNumber(vTotalRecebido - vTotalSaidas, 2)
    Else
        lblVlrRestante2.Caption = FormatNumber(0, 2)
    End If
    'lblVlrRestante.Caption = FormatNumber(vRestante, 2)
Else
    lblVlrRestante.Caption = FormatNumber(0, 2)
    lblVlrRestante2.Caption = FormatNumber(0, 2)
End If

If r.State <> 0 Then r.Close
Set r = Nothing


'valor NFCE
'Dim vTotalNFCE As Currency
'sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.BOF Then
'    lblNFCe.Caption = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
'    vTotalNFCE = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
'Else
'    lblNFCe.Caption = FormatNumber(0, 2)
'    vTotalNFCE = FormatNumber(0, 2)
'End If

''valor NFE
'Dim vTotalNFe As Currency
'sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.BOF Then
'    lblNFe.Caption = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
'    vTotalNFe = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
'Else
'    lblNFe.Caption = FormatNumber(0, 2)
'    vTotalNFe = FormatNumber(0, 2)
'End If

'somar totais
'Dim vTotalEntradas As Currency
'Dim vTotalSaidas As Currency
'Dim vRestante As Currency

'vTotalEntradas = vTotalXML + vTotalManual
'vTotalSaidas = vTotalNFCE + vTotalNFe

'lblVlrEntrada.Caption = FormatNumber(vTotalEntradas, 2)
'lblVlrSaida.Caption = FormatNumber(vTotalSaidas, 2)

'If vTotalRecebido > vTotalEntradas Then
'    lblMeta.Caption = FormatNumber(vTotalRecebido, 2)
'ElseIf vTotalRecebido < vTotalEntradas Then
'    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
'ElseIf vTotalRecebido = vTotalEntradas Then
'    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
'End If
'Dim vMeta As Currency
'vMeta = FormatNumber(lblMeta.Caption, 2)

'lblEmitido.Caption = FormatNumber(vTotalSaidas, 2)

'vRestante = vMeta - vTotalSaidas

'If vRestante > 0 Then
'    lblVlrRestante.Caption = FormatNumber(vRestante, 2)
'Else
'    lblVlrRestante.Caption = FormatNumber(0, 2)
'End If

'If r.State <> 0 Then r.Close
'Set r = Nothing

End Sub

Private Sub lblVoltar_Click()
'definição de mês
vDataHoje = Format(DateAdd("m", -Val(1), vDataHoje), "dd/mm/yy")

vMesInt = Format(vDataHoje, "mmmm")
vAno = Year(vDataHoje)
lblMesRef.Caption = vMesInt & "/" & vAno


If vMesInt = "janeiro" Then
    vNumeroMes = 1
ElseIf vMesInt = "fevereiro" Then
    vNumeroMes = 2
ElseIf vMesInt = "março" Then
    vNumeroMes = 3
ElseIf vMesInt = "abril" Then
    vNumeroMes = 4
ElseIf vMesInt = "maio" Then
    vNumeroMes = 5
ElseIf vMesInt = "junho" Then
    vNumeroMes = 6
ElseIf vMesInt = "julho" Then
    vNumeroMes = 7
ElseIf vMesInt = "agosto" Then
    vNumeroMes = 8
ElseIf vMesInt = "setembro" Then
    vNumeroMes = 9
ElseIf vMesInt = "outubro" Then
    vNumeroMes = 10
ElseIf vMesInt = "novembro" Then
    vNumeroMes = 11
ElseIf vMesInt = "dezembro" Then
    vNumeroMes = 12
End If

Mostrar_RecPIX
Mostrar_RecCartaoC
Mostrar_RecCartaoD
Mostrar_RecDeposito
Mostrar_RecTransf

Dim vTotalRecebido As Currency
vTotalRecebido = vTotalPix + vTotalCartaoC + vTotalCartaoD + vTotalTransf + vTotalDeposito
lblTotalRecebido.Caption = FormatNumber(vTotalRecebido, 2)

Mostrar_Entradas

'valor NFCE
Dim vTotalNFCE As Currency
If optMensal.Value = True Then
    sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (YEAR(DataEmissao) = " & cboAno & ")"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblNFCe.Caption = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
    vTotalNFCE = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
Else
    lblNFCe.Caption = FormatNumber(0, 2)
    vTotalNFCE = FormatNumber(0, 2)
End If

'valor NFE
Dim vTotalNFe As Currency
If optMensal.Value = True Then
    sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
Else
    If cboAno.Text = "" Then cboAno.Text = vAno
    sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (YEAR(DataEmissao) = " & cboAno & ")"
End If
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblNFe.Caption = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
    vTotalNFe = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
Else
    lblNFe.Caption = FormatNumber(0, 2)
    vTotalNFe = FormatNumber(0, 2)
End If

'somar totais
Dim vTotalEntradas As Currency
Dim vTotalSaidas As Currency
Dim vRestante As Currency

vTotalEntradas = vTotalXML + vTotalManual
vTotalSaidas = vTotalNFCE + vTotalNFe

lblVlrEntrada.Caption = FormatNumber(vTotalEntradas, 2)
lblVlrSaida.Caption = FormatNumber(vTotalSaidas, 2)
lblVlrSaida2.Caption = FormatNumber(vTotalSaidas, 2)

vTotalEntradas = vTotalEntradas + (vTotalEntradas * 0.25)
lblMeta1.Caption = FormatNumber(vTotalEntradas, 2)  'META 1

If vTotalRecebido > vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalRecebido, 2)
ElseIf vTotalRecebido < vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
ElseIf vTotalRecebido = vTotalEntradas Then
    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
End If
Dim vMeta As Currency
vMeta = FormatNumber(lblMeta.Caption, 2)            'META 3

lblEmitido.Caption = FormatNumber(vTotalRecebido, 2)  'META 2

vRestante = vMeta - vTotalSaidas

If vRestante > 0 Then
    lblVlrRestante.Caption = FormatNumber(vTotalEntradas - vTotalSaidas, 2)
    If vTotalRecebido > 0 Then
        lblVlrRestante2.Caption = FormatNumber(vTotalRecebido - vTotalSaidas, 2)
    Else
        lblVlrRestante2.Caption = FormatNumber(0, 2)
    End If
    'lblVlrRestante.Caption = FormatNumber(vRestante, 2)
Else
    lblVlrRestante.Caption = FormatNumber(0, 2)
    lblVlrRestante2.Caption = FormatNumber(0, 2)
End If

If r.State <> 0 Then r.Close
Set r = Nothing

''valor NFCE
'Dim vTotalNFCE As Currency
'sSQL = "SELECT SUM(Valor_NF_Prod - DescontoPromocional) AS vTotalNFCe FROM  TbNFCe WHERE (NFCeEnviada = 1) AND (NFCeCancelada = 0) AND (Inutilizada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.BOF Then
'    lblNFCe.Caption = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
'    vTotalNFCE = FormatNumber(ValidateNull(r("vTotalNFCe")), 2)
'Else
'    lblNFCe.Caption = FormatNumber(0, 2)
'    vTotalNFCE = FormatNumber(0, 2)
'End If

''valor NFE
'Dim vTotalNFe As Currency
'sSQL = "SELECT SUM(valornota) AS vTotalNFe FROM  NotaFiscal WHERE (Enviada = 1) AND (Cancelada = 0) AND (MONTH(DataEmissao) = " & vNumeroMes & ") AND (YEAR(DataEmissao) = " & vAno & ")"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.BOF Then
'    lblNFe.Caption = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
'    vTotalNFe = FormatNumber(ValidateNull(r("vTotalNFe")), 2)
'Else
'    lblNFe.Caption = FormatNumber(0, 2)
'    vTotalNFe = FormatNumber(0, 2)
'End If

''somar totais
'Dim vTotalEntradas As Currency
'Dim vTotalSaidas As Currency
'Dim vRestante As Currency

'vTotalEntradas = vTotalXML + vTotalManual
'vTotalSaidas = vTotalNFCE + vTotalNFe

'lblVlrEntrada.Caption = FormatNumber(vTotalEntradas, 2)
'lblVlrSaida.Caption = FormatNumber(vTotalSaidas, 2)

'If vTotalRecebido > vTotalEntradas Then
'    lblMeta.Caption = FormatNumber(vTotalRecebido, 2)
'ElseIf vTotalRecebido < vTotalEntradas Then
'    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
'ElseIf vTotalRecebido = vTotalEntradas Then
'    lblMeta.Caption = FormatNumber(vTotalEntradas, 2)
'End If
'Dim vMeta As Currency
'vMeta = FormatNumber(lblMeta.Caption, 2)

'lblEmitido.Caption = FormatNumber(vTotalSaidas, 2)

'vRestante = vMeta - vTotalSaidas

'If vRestante > 0 Then
'    lblVlrRestante.Caption = FormatNumber(vRestante, 2)
'Else
'    lblVlrRestante.Caption = FormatNumber(0, 2)
'End If

'If r.State <> 0 Then r.Close
'Set r = Nothing

End Sub


Private Sub optAnual_Click()
Form_Load
lblVoltar.Visible = False
Label14.Visible = False
lblAvancar.Visible = False
lblMesRef.Visible = False

lblAno.Visible = True
cboAno.Visible = True
End Sub

Private Sub optMensal_Click()
Form_Load

lblVoltar.Visible = True
Label14.Visible = True
lblAvancar.Visible = True
lblMesRef.Visible = True

lblAno.Visible = False
cboAno.Visible = False
End Sub


Private Sub Timer1_Timer()

    ' --- Condição 1: Total Recebido ---
    If lblTotalRecebido.Caption = "0,00" Then
        imgExcl1.Visible = Not imgExcl1.Visible
    Else
        imgExcl1.Visible = False ' Fica invisível se o valor for preenchido
    End If

    ' --- Condição 2: Valor Entrada ---
    If lblVlrEntrada.Caption = "0,00" Then
        imgExcl2.Visible = Not imgExcl2.Visible
    Else
        imgExcl2.Visible = False
    End If

    ' --- Condição 3: Valor Saída ---
    If lblVlrSaida.Caption = "0,00" Then
        imgExcl3.Visible = Not imgExcl3.Visible
    Else
        imgExcl3.Visible = False
    End If

    ' --- Condição 4: Valor Restante (Um ou Outro) ---
    If lblVlrRestante.Caption = "0,00" Or lblVlrRestante2.Caption = "0,00" Then
        imgExcl4.Visible = Not imgExcl4.Visible
    Else
        imgExcl4.Visible = False
    End If
End Sub


