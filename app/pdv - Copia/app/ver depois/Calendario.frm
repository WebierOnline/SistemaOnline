VERSION 5.00
Object = "{703944EE-9203-11D2-8865-AD1268A0A52F}#1.0#0"; "ActiveCal.ocx"
Begin VB.Form Calendario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendário"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "Calendario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin rdActiveCal.ActiveCalendar ActiveCalendar1 
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   6244
      Date            =   36623
      TodayCaption    =   "&Today"
      BorderStyle     =   0
   End
   Begin VB.Label lblObjeto 
      BackColor       =   &H80000009&
      Height          =   195
      Left            =   2640
      TabIndex        =   1
      Top             =   3550
      Width           =   915
   End
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dataSelecionada As Date

Public Property Get DateSelected() As Date
   DateSelected = dataSelecionada
End Property
Private Sub ActiveCalendar1_DblClick()
   dataSelecionada = ActiveCalendar1.SelectedDate
   Unload Me
End Sub

Private Sub Form_Resize()
    ActiveCalendar1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
