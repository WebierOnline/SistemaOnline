VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Consulta_Estoque_Minimo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ESTOQUE MÍNIMO"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "Estoque_Minimo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMostrarZero 
      Caption         =   "Mostrar produtos com estoque minimo zerado."
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
      TabIndex        =   4
      Top             =   6840
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5595
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   9869
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   9405
      TabIndex        =   0
      Top             =   60
      Width           =   9435
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTOQUE MÍNIMO"
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
         Left            =   1365
         TabIndex        =   1
         Top             =   240
         Width           =   2715
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Estoque_Minimo.frx":23D2
         Top             =   120
         Width           =   645
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   555
      Left            =   7800
      TabIndex        =   2
      Top             =   6660
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Estoque_Minimo.frx":7DA5
      PICN            =   "Estoque_Minimo.frx":7DC1
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
      TabIndex        =   5
      Top             =   7275
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12594
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:30"
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
Attribute VB_Name = "Consulta_Estoque_Minimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Visible = False
      
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 5800
      .ColWidth(3) = 1600
      .ColWidth(4) = 1600
      
      '.RowHeight(0) = 0
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "QUANT. ATUAL"
      .TextMatrix(0, 4) = "QUANT. MÍNIMA"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("descricao"))
            .TextMatrix(.Rows - 1, 3) = rTabela("quant_estoque")
            .TextMatrix(.Rows - 1, 4) = rTabela("quant_min")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      .Visible = True
   End With
End Sub

Private Sub chkMostrarZero_Click()
   Form_Load
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    Dim r As ADODB.Recordset
    
   If chkMostrarZero.Value = Checked Then
      sSQL = "SELECT codigo, descricao, quant_estoque, quant_min FROM produtos WHERE (quant_estoque < quant_min) OR (quant_min = 0) ORDER BY descricao;"
    
   Else
      sSQL = "SELECT codigo, descricao, quant_estoque, quant_min FROM produtos WHERE (quant_estoque < quant_min) ORDER BY descricao;"
   
   End If
   
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
