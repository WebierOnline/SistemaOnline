VERSION 5.00
Begin VB.Form Notas_Adesivas 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RESUMO FISCAL"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblVoltar 
      Alignment       =   2
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Mês anterior"
      Top             =   90
      Width           =   480
      BackColor       =   &H00FFFFFF&
   End
   Begin VB.Label lblMesRef 
      Alignment       =   2
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   660
      TabIndex        =   1
      Top             =   150
      Width           =   3285
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblAvancar 
      Alignment       =   2
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   4005
      TabIndex        =   2
      ToolTipText     =   "Próximo mês"
      Top             =   90
      Width           =   480
      BackColor       =   &H00FFFFFF&
   End
   Begin VB.Label lblAno 
      Alignment       =   0
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1200
      TabIndex        =   3
      Top             =   150
      Width           =   600
      BackStyle       =   0  'Transparent
   End
   Begin VB.ComboBox cboAno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.OptionButton optAnual 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver o ano todo"
      Height          =   225
      Left            =   2400
      TabIndex        =   5
      Top             =   570
      Width           =   1700
   End
   Begin VB.OptionButton optMensal 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver um mes"
      Height          =   225
      Left            =   600
      TabIndex        =   6
      Top             =   570
      Value           =   -1  'True
      Width           =   1700
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2
      Caption         =   "Falta emitir de nota fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   870
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblFalta 
      Alignment       =   2
      Caption         =   "R$ 0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Top             =   1100
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblSituacao 
      Alignment       =   2
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   1520
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblBarraFundo 
      Alignment       =   0
      BorderStyle     =   1  'Fixed Single
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4365
      BackColor       =   &H00FFFFFF&
   End
   Begin VB.Label lblBarra 
      Alignment       =   0
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   135
      TabIndex        =   11
      Top             =   1815
      Width           =   15
      BackColor       =   &H00808080&
   End
   Begin VB.Label Label1 
      Alignment       =   0
      Caption         =   "O QUE COMPREI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Left            =   120
      TabIndex        =   12
      Top             =   2180
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Line Line101 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4485
      Y1              =   2435
      Y2              =   2435
   End
   Begin VB.Label Label2 
      Alignment       =   0
      Caption         =   "Total das notas de compra (XML):"
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
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Soma das notas fiscais de compra importadas por XML no período"
      Top             =   2455
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblVlrEntrada 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Soma das notas fiscais de compra importadas por XML no período"
      Top             =   2455
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label3 
      Alignment       =   0
      Caption         =   "O QUE RECEBI DOS CLIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Left            =   120
      TabIndex        =   15
      Top             =   2780
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Line Line103 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4485
      Y1              =   3035
      Y2              =   3035
   End
   Begin VB.Label Label4 
      Alignment       =   0
      Caption         =   "Pix:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   16
      Top             =   3055
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblRecPix 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   3055
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label5 
      Alignment       =   0
      Caption         =   "Cartão de débito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Top             =   3280
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblRecDebito 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   19
      Top             =   3280
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label6 
      Alignment       =   0
      Caption         =   "Cartão de crédito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   20
      Top             =   3505
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblRecCredito 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   3505
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label7 
      Alignment       =   0
      Caption         =   "Transferência:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   22
      Top             =   3730
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblRecTransf 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   23
      Top             =   3730
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label8 
      Alignment       =   0
      Caption         =   "Depósito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   24
      Top             =   3955
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblRecDepos 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   25
      Top             =   3955
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Line LineRec 
      BorderColor     =   &H00808080&
      X1              =   3000
      X2              =   4485
      Y1              =   4210
      Y2              =   4210
   End
   Begin VB.Label Label9 
      Alignment       =   0
      Caption         =   "Total recebido:"
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
      Left            =   120
      TabIndex        =   26
      Top             =   4250
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblTotalRecebido 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   27
      Top             =   4250
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label10 
      Alignment       =   0
      Caption         =   "O QUE EMITI DE NOTA (vendas)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Left            =   120
      TabIndex        =   28
      Top             =   4580
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Line Line110 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4485
      Y1              =   4835
      Y2              =   4835
   End
   Begin VB.Label Label11 
      Alignment       =   0
      Caption         =   "Notas NF-e:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "NF-e de saída enviadas, sem canceladas, devoluções e ajustes"
      Top             =   4855
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblNFe 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   30
      ToolTipText     =   "NF-e de saída enviadas, sem canceladas, devoluções e ajustes"
      Top             =   4855
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label12 
      Alignment       =   0
      Caption         =   "Notas NFC-e:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   31
      ToolTipText     =   "NFC-e enviadas, sem canceladas e inutilizadas"
      Top             =   5080
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblNFCe 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   32
      ToolTipText     =   "NFC-e enviadas, sem canceladas e inutilizadas"
      Top             =   5080
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Line LineEmi 
      BorderColor     =   &H00808080&
      X1              =   3000
      X2              =   4485
      Y1              =   5335
      Y2              =   5335
   End
   Begin VB.Label Label13 
      Alignment       =   0
      Caption         =   "Total emitido:"
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
      Left            =   120
      TabIndex        =   33
      Top             =   5375
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblVlrSaida 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   34
      Top             =   5375
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label14 
      Alignment       =   0
      Caption         =   "METAS DO PERÍODO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Left            =   120
      TabIndex        =   35
      Top             =   5700
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.Line Line114 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4485
      Y1              =   5955
      Y2              =   5955
   End
   Begin VB.Label lblTxtMinimo 
      Alignment       =   0
      Caption         =   "Mínimo recomendado (+ 25%):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   36
      ToolTipText     =   "O que você comprou + 25%. É o mínimo que os contadores pedem para emitir"
      Top             =   5975
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblMeta1 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   37
      ToolTipText     =   "O que você comprou + 25%. É o mínimo que os contadores pedem para emitir"
      Top             =   5975
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label15 
      Alignment       =   0
      Caption         =   "Meta do mês (a maior):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   38
      ToolTipText     =   "O maior valor entre o mínimo recomendado e o que você recebeu"
      Top             =   6200
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblMeta 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   39
      ToolTipText     =   "O maior valor entre o mínimo recomendado e o que você recebeu"
      Top             =   6200
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label16 
      Alignment       =   0
      Caption         =   "Falta emitir para o mínimo:"
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
      Left            =   120
      TabIndex        =   40
      Top             =   6425
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblVlrRestante 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   41
      Top             =   6425
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label Label17 
      Alignment       =   0
      Caption         =   "Falta para igualar o que recebi:"
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
      Left            =   120
      TabIndex        =   42
      Top             =   6650
      Width           =   2900
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblVlrRestante2 
      Alignment       =   1
      Caption         =   "0,00"
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
      Left            =   3000
      TabIndex        =   43
      Top             =   6650
      Width           =   1485
      BackStyle       =   0  'Transparent
   End
   Begin VB.Label lblAjuda 
      Alignment       =   2
      Caption         =   "Como funciona esta tela? (clique aqui)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   44
      ToolTipText     =   "Clique para ver a explicação"
      Top             =   6930
      Width           =   4365
      BackStyle       =   0  'Transparent
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir / PDF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   45
      ToolTipText     =   "Gera o relatório do resumo (na Impressora PDF para salvar em PDF)"
      Top             =   7230
      Width           =   2130
   End
   Begin VB.CommandButton cmdEnviarContab 
      Caption         =   "Enviar à contabilidade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2355
      TabIndex        =   46
      ToolTipText     =   "Gera o PDF e envia por e-mail para a contabilidade"
      Top             =   7230
      Width           =   2130
   End
End
Attribute VB_Name = "Notas_Adesivas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vDataHoje As Date           'mes de referencia (modo Mensal)
Dim bCarregando As Boolean      'bloqueia recalculo enquanto o form monta os controles
Dim vAnoCalc As Integer         'ultimo ano (modo Anual) usado no calculo
Dim rfAtual As TResumoFiscal    'ultimo resumo calculado (usado no relatorio e no envio) - ver modResumoFiscal.bas

Private Sub Form_Load()
Dim i As Integer

bCarregando = True
vDataHoje = Date

lblTxtMinimo.Caption = "Mínimo recomendado (+ " & Format(ACRESCIMO_META * 100, "0") & "%):"

cboAno.Clear
For i = Year(Date) - 10 To Year(Date)
   cboAno.AddItem CStr(i)
Next i
cboAno.Text = CStr(Year(Date))
vAnoCalc = Year(Date)

lblBarra.ZOrder 0   'preenchimento da barra por cima do fundo

bCarregando = False
AjustarModo
AtualizarResumo
End Sub

'Mostra os controles do modo escolhido (Mensal = navegacao por mes; Anual = combo de ano)
Private Sub AjustarModo()
Dim bMensal As Boolean

bMensal = optMensal.Value
lblVoltar.Visible = bMensal
lblAvancar.Visible = bMensal
lblMesRef.Visible = bMensal
lblAno.Visible = Not bMensal
cboAno.Visible = Not bMensal
End Sub

Private Sub AtualizarResumo()
Dim nAno As Integer

If bCarregando Then Exit Sub

If optMensal.Value Then nAno = Year(vDataHoje) Else nAno = vAnoCalc
rfAtual = RF_Calcular(optMensal.Value, Month(vDataHoje), nAno)

If optMensal.Value Then lblMesRef.Caption = rfAtual.sPeriodo

lblRecPix.Caption = Format(rfAtual.vPix, ocMONEY)
lblRecCredito.Caption = Format(rfAtual.vCredito, ocMONEY)
lblRecDebito.Caption = Format(rfAtual.vDebito, ocMONEY)
lblRecTransf.Caption = Format(rfAtual.vTransf, ocMONEY)
lblRecDepos.Caption = Format(rfAtual.vDepos, ocMONEY)
lblTotalRecebido.Caption = Format(rfAtual.vTotalRec, ocMONEY)
lblVlrEntrada.Caption = Format(rfAtual.vEntradas, ocMONEY)
lblNFCe.Caption = Format(rfAtual.vNFCe, ocMONEY)
lblNFe.Caption = Format(rfAtual.vNFe, ocMONEY)
lblVlrSaida.Caption = Format(rfAtual.vTotalSaidas, ocMONEY)
lblMeta1.Caption = Format(rfAtual.vMeta1, ocMONEY)          'META 1
lblMeta.Caption = Format(rfAtual.vMeta3, ocMONEY)           'META 3
lblVlrRestante.Caption = Format(rfAtual.vRestante1, ocMONEY)
lblVlrRestante2.Caption = Format(rfAtual.vRestante2, ocMONEY)

MostrarSituacao rfAtual
End Sub

'Frase de destaque + barra de progresso (verde = meta atingida, laranja = perto, vermelho = longe)
Private Sub MostrarSituacao(ByRef r As TResumoFiscal)
Dim dPct As Double
Dim lCor As Long
Dim lLarg As Long

lblTitulo.Caption = "Falta emitir em " & r.sPeriodo & ":"
lblFalta.Caption = "R$ " & Format(r.vRestante1, ocMONEY)
lblSituacao.Caption = RF_Situacao(r, dPct)

If r.vMeta1 <= 0 Then
   lCor = &H808080
ElseIf dPct >= 1 Then
   lCor = &H8000&
ElseIf dPct >= 0.6 Then
   lCor = &H80FF&
Else
   lCor = &HC0&
End If
If dPct > 1 Then dPct = 1

lLarg = CLng((lblBarraFundo.Width - 30) * dPct)
If lLarg < 15 Then lLarg = 15
lblBarra.Width = lLarg
lblBarra.Visible = (dPct > 0)
lblBarra.BackColor = lCor
lblFalta.ForeColor = lCor
lblSituacao.ForeColor = lCor
End Sub

Private Sub lblAjuda_Click()
MsgBox "COMO LER ESTA TELA" & vbCrLf & vbCrLf & _
       "O que comprei: total das notas de compra (XML) do período." & vbCrLf & vbCrLf & _
       "O que recebi dos clientes: dinheiro recebido por Pix, cartão, transferência e depósito." & vbCrLf & vbCrLf & _
       "O que emiti de nota: notas fiscais de venda (NF-e e NFC-e) enviadas, sem as canceladas, devoluções e ajustes." & vbCrLf & vbCrLf & _
       "Mínimo recomendado: o que você comprou + " & Format(ACRESCIMO_META * 100, "0") & "%. Os contadores pedem para emitir nota de pelo menos " & _
       Format(ACRESCIMO_META * 100, "0") & "% a mais do que o valor das compras." & vbCrLf & vbCrLf & _
       "Meta do mês: o maior valor entre o mínimo recomendado e o que você recebeu dos clientes." & vbCrLf & vbCrLf & _
       "A frase grande no topo diz quanto ainda falta emitir para chegar no mínimo recomendado.", vbInformation, "Resumo Fiscal"
End Sub

Private Sub lblAvancar_Click()
vDataHoje = DateAdd("m", 1, vDataHoje)
AtualizarResumo
End Sub

Private Sub lblVoltar_Click()
vDataHoje = DateAdd("m", -1, vDataHoje)
AtualizarResumo
End Sub

Private Sub optMensal_Click()
AjustarModo
AtualizarResumo
End Sub

Private Sub optAnual_Click()
AjustarModo
AtualizarResumo
End Sub

Private Sub cboAno_Click()
AplicarAno
End Sub

Private Sub cboAno_LostFocus()
AplicarAno
End Sub

'Valida o ano do combo e recalcula so quando o ano mudou
Private Sub AplicarAno()
Dim vAno As Integer

If bCarregando Then Exit Sub

If cboAno.ListIndex >= 0 Then
   vAno = Val(cboAno.List(cboAno.ListIndex))
Else
   vAno = Val(cboAno.Text)
End If

If Len(Trim(cboAno.Text)) <> 4 Or vAno < 1900 Or vAno > 2100 Then
   MsgBox "Informe um ano válido com 4 dígitos (ex: " & Year(Date) & ").", vbExclamation, "Ano Inválido"
   cboAno.Text = CStr(vAnoCalc)
   Exit Sub
End If

If vAno <> vAnoCalc Then
   vAnoCalc = vAno
   AtualizarResumo
End If
End Sub


'---------------------------------------------------------------------------
' Relatorio (REL_ResumoFiscal) e envio para a contabilidade
' (calculo, PDF e e-mail ficam em modResumoFiscal.bas, tambem usado no envio automatico)
'---------------------------------------------------------------------------
Private Function ExisteImpressoraPDF() As Boolean
Dim Prt As Printer

For Each Prt In Printers
   If Prt.DeviceName = "Impressora PDF" Then
      ExisteImpressoraPDF = True
      Exit For
   End If
Next
End Function

'Gera o relatorio na impressora informada; com sImpressora vazio abre so a visualizacao
Private Sub EmitirRelatorio(ByVal sImpressora As String)
Dim i As Long

Screen.MousePointer = vbHourglass
RF_MontarLinhas rfAtual
REL_ResumoFiscal.LimparLinhas
For i = 1 To RF_LinhasContar
   REL_ResumoFiscal.AdicionarLinha RF_LinhaTexto(i), RF_LinhaValor(i), RF_LinhaNegrito(i), RF_LinhaTamanho(i), RF_LinhaCentro(i)
Next

With REL_ResumoFiscal.ReportMain1
   .NumeroRegistros = REL_ResumoFiscal.ContagemLinhas
   .Titulo = "Resumo Fiscal - " & rfAtual.sPeriodo
   If sImpressora <> "" Then
      .NomeImpressora = sImpressora
      .Visualizar = False
   Else
      .Visualizar = True
   End If
   Screen.MousePointer = vbDefault
   .Ativar
End With
Unload REL_ResumoFiscal
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdImprimir_Click()
If ExisteImpressoraPDF Then
   EmitirRelatorio "Impressora PDF"
Else
   EmitirRelatorio ""
End If
End Sub

'Gera o PDF na pasta RESUMOS e envia por e-mail para a contabilidade
Private Sub cmdEnviarContab_Click()
Dim sEmail As String
Dim sArquivo As String
Dim nAno As Integer

'e-mail da contabilidade (cadastro do contador); se nao tiver, pergunta
sEmail = RF_EmailContador
If sEmail = "" Then sEmail = Trim(InputBox("Informe o e-mail da contabilidade", "Envio de Email", ""))
If sEmail = "" Then Exit Sub
If InStr(1, sEmail, "@") = 0 Then
   MsgBox "E-mail inválido! O endereço não contém '@'.", vbExclamation, "Aviso"
   Exit Sub
End If

If MsgBox("Gerar o PDF do Resumo Fiscal de " & rfAtual.sPeriodo & " e enviar para:" & vbCrLf & sEmail, _
          vbOKCancel + vbQuestion, "Enviar à contabilidade") = vbCancel Then Exit Sub

If optMensal.Value Then nAno = Year(vDataHoje) Else nAno = vAnoCalc

Screen.MousePointer = vbHourglass
sArquivo = RF_GerarPDF(rfAtual, optMensal.Value, Month(vDataHoje), nAno)
If sArquivo = "" Then
   Screen.MousePointer = vbDefault
   MsgBox RF_UltimoErro, vbCritical, "Aviso do Sistema"
   Exit Sub
End If

If RF_EnviarEmail(sEmail, sArquivo, rfAtual.sPeriodo, False) Then
   Screen.MousePointer = vbDefault
   MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"
Else
   Screen.MousePointer = vbDefault
   MsgBox RF_UltimoErro, vbCritical + vbOKOnly, "ERRO: Envio Email"
End If
End Sub
