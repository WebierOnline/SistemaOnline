VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Aniversariantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANIVERSÁRIOS"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "Aniversariantes_Clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   2640
      ScaleHeight     =   795
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   840
      Width           =   2715
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   375
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "MÊS :"
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
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   5265
      TabIndex        =   2
      Top             =   60
      Width           =   5295
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ANIVERSÁRIOS"
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
         Left            =   960
         TabIndex        =   3
         Top             =   180
         Width           =   2340
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   180
         Picture         =   "Aniversariantes_Clientes.frx":23D2
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   60
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   840
      Width           =   2535
      Begin VB.OptionButton optClientes 
         Caption         =   "&Clientes"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   1755
      End
      Begin VB.OptionButton optFuncionarios 
         Caption         =   "&Funcionários"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdGerar 
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   1800
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
      MICON           =   "Aniversariantes_Clientes.frx":8B19
      PICN            =   "Aniversariantes_Clientes.frx":8B35
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
      TabIndex        =   9
      Top             =   2505
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5239
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "07:56"
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
Attribute VB_Name = "Aniversariantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGerar_Click()
   'colocar o nome da maquina na barra de status
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Dim sSQL As String
   Dim r As ADODB.Recordset, r2 As ADODB.Recordset
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   If optClientes.Value = True Then
      sSQL = "SELECT nome, cidade, data_de_nascimento AS var_nasc, CELULAR FROM cliente WHERE (MONTH(data_de_nascimento) = " & cboMes.ListIndex + 1 & ") ORDER BY DAY(data_de_nascimento);"
      
   ElseIf optFuncionarios.Value = True Then
      sSQL = "SELECT nome, cidade, nascimento AS var_nasc, CELULAR FROM funcionario WHERE (MONTH(nascimento) = " & cboMes.ListIndex + 1 & ") ORDER BY DAY(nascimento);"
   
   End If
   
   Set r = dbData.OpenRecordset(sSQL)
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r2 = dbData.OpenRecordset(sSQL)
   
   Me.Hide
   'Principal_Impressao.Hide
   
   Set REL_Aniversariantes.Relatorio.Recordset = r
   REL_Aniversariantes.ReportField1.Caption = REL_Aniversariantes.ReportField1.Caption & " " & cboMes.Text
   REL_Aniversariantes.Label1.Caption = "A " & r2("fantasia") & " deseja a todos um feliz anivesário."
   'REL_Aniversariantes.Relatorio.NomeImpressora = var_Impressora
   REL_Aniversariantes.Relatorio.Ativar
   Unload REL_Aniversariantes
   
   Me.Show
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   For i = 1 To 12
      cboMes.AddItem UCase(MonthName(i))
   Next
   
   cboMes.ListIndex = Month(Date) - 1
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub

