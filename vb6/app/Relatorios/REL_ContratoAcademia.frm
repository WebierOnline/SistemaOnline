VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form REL_ContratoAcademia 
   Caption         =   "Form1"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   204.523
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   224.631
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain ReportMain1 
      Height          =   480
      Left            =   12180
      TabIndex        =   0
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   11535
      Left            =   0
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   20346
      AutoExpandir    =   -1  'True
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "---"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   10080
         Width           =   11715
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(Nome e assinatura do Representante legal da Contratada)"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   9720
         Width           =   11715
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(Nome e assinatura do Contratante)"
         Height          =   315
         Left            =   0
         TabIndex        =   23
         Top             =   9420
         Width           =   11715
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RESCISÃO"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   11715
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":0000
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   8640
         Width           =   11715
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cláusula 25ª. Para dirimir quaisquer controvérsias oriundas do CONTRATO, as partes elegem o foro da comarca de (xxx);"
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   8340
         Width           =   11715
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DO FORO"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   8040
         Width           =   11715
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":008B
         Height          =   555
         Left            =   120
         TabIndex        =   18
         Top             =   7440
         Width           =   11715
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":0161
         Height          =   555
         Left            =   60
         TabIndex        =   17
         Top             =   6840
         Width           =   11715
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RESCISÃO"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   6540
         Width           =   11715
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cláusula 21ª. O presente contrato será de prazo indeterminado. "
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   6180
         Width           =   11715
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DO PRAZO"
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   5820
         Width           =   11715
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cláusula 13ª. O CONTRATANTE pagará o valor mensal de R$ (xxx) (Valor expresso), todo dia (xxx) de cada mês."
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   5400
         Width           =   11715
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":0288
         Height          =   555
         Left            =   60
         TabIndex        =   11
         Top             =   4800
         Width           =   11715
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":033B
         Height          =   555
         Left            =   60
         TabIndex        =   10
         Top             =   4200
         Width           =   11715
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":03FB
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   3900
         Width           =   11715
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAGAMENTO DAS MENSALIDADES"
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   3540
         Width           =   11715
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cláusula 6ª. A CONTRATADA terá técnicos qualificados para orientação do CONTRATANTE."
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   3180
         Width           =   11715
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":0491
         Height          =   555
         Left            =   60
         TabIndex        =   6
         Top             =   2580
         Width           =   11715
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DO OBJETO DO CONTRATO"
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   2220
         Width           =   11715
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":0566
         Height          =   555
         Left            =   60
         TabIndex        =   4
         Top             =   1620
         Width           =   11715
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":063F
         Height          =   555
         Left            =   60
         TabIndex        =   3
         Top             =   1020
         Width           =   11715
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"REL_ContratoAcademia.frx":084E
         Height          =   555
         Left            =   60
         TabIndex        =   2
         Top             =   420
         Width           =   11715
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(Local, data e ano)."
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   9000
         Width           =   11715
      End
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DO PRAZO"
      Height          =   555
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11715
   End
End
Attribute VB_Name = "REL_ContratoAcademia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

