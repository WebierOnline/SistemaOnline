VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Principal_Impressao 
   Caption         =   "IMPRESSÕES"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "Principal_Impressão.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4755
      Left            =   2580
      ScaleHeight     =   4695
      ScaleWidth      =   2415
      TabIndex        =   3
      Top             =   1020
      Width           =   2475
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   4965
      TabIndex        =   1
      Top             =   0
      Width           =   4995
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESSÕES"
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
         Left            =   45
         TabIndex        =   2
         Top             =   300
         Width           =   4830
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   360
         Picture         =   "Principal_Impressão.frx":23D2
         Top             =   60
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4755
      Left            =   60
      ScaleHeight     =   4695
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   1020
      Width           =   2475
      Begin ChamaleonBtn.chameleonButton cmdIMPRecibo 
         Height          =   615
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Principal_Impressão.frx":8D04
         PICN            =   "Principal_Impressão.frx":8D20
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdIMPPromissoria 
         Height          =   615
         Left            =   60
         TabIndex        =   5
         Top             =   1380
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Principal_Impressão.frx":E914
         PICN            =   "Principal_Impressão.frx":E930
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdIMPReciboAvulso 
         Height          =   615
         Left            =   60
         TabIndex        =   6
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Principal_Impressão.frx":149C4
         PICN            =   "Principal_Impressão.frx":149E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAniversario 
         Height          =   615
         Left            =   60
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Principal_Impressão.frx":1583B
         PICN            =   "Principal_Impressão.frx":15857
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   5880
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4233
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "19:01"
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
Attribute VB_Name = "Principal_Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAniversario_Click()
   On Error GoTo Erro
   Unload Me
   Aniversariantes.Show 1
   Exit Sub
   
Erro:
   If Err.Number = 400 Then Unload Aniversariantes
End Sub

Private Sub cmdIMPRecibo_Click()
   Unload Me
   Recibo.Show 1
End Sub

Private Sub cmdIMPReciboAvulso_Click()
   Unload Me
   Recibos_Avulso.Show 1
End Sub

Private Sub cmdIMPPromissoria_Click()
   Unload Me
   Promissoria.Show 1
End Sub

Private Sub Form_Load()
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub
