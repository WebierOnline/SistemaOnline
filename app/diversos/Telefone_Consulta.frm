VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Telefone_Consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TELEFONES"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   Icon            =   "Telefone_Consulta.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   60
      ScaleHeight     =   1665
      ScaleWidth      =   13365
      TabIndex        =   4
      Top             =   60
      Width           =   13395
      Begin VB.Image Image1 
         Height          =   1500
         Left            =   60
         Picture         =   "Telefone_Consulta.frx":030A
         Top             =   60
         Width           =   6450
      End
      Begin VB.Image Image2 
         Height          =   1590
         Left            =   8820
         MousePointer    =   14  'Arrow and Question
         Picture         =   "Telefone_Consulta.frx":39DE
         Top             =   60
         Width           =   4500
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5415
      Left            =   60
      TabIndex        =   3
      Top             =   1920
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   9551
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busca Avançada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4200
      TabIndex        =   1
      Top             =   7440
      Width           =   5445
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5175
      End
   End
   Begin ChamaleonBtn.chameleonButton Command6 
      Height          =   675
      Left            =   11700
      TabIndex        =   2
      Top             =   7440
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1191
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
      MICON           =   "Telefone_Consulta.frx":747B
      PICN            =   "Telefone_Consulta.frx":7497
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
      Top             =   8220
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19526
            Text            =   "Seja bem-vindo..."
            TextSave        =   "Seja bem-vindo..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:48"
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
   Begin ChamaleonBtn.chameleonButton Command4 
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1191
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
      MICON           =   "Telefone_Consulta.frx":77B1
      PICN            =   "Telefone_Consulta.frx":77CD
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
Attribute VB_Name = "Telefone_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Private Sub Atualizar_Grid()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM telefone ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGrid r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 11
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 3400
      .ColWidth(3) = 1300
      .ColWidth(4) = 1300
      .ColWidth(5) = 1300
      .ColWidth(6) = 1300
      .ColWidth(7) = 900
      .ColWidth(8) = 1300
      .ColWidth(9) = 900
      .ColWidth(10) = 1300
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "NOME"
      .TextMatrix(0, 3) = "PESSOAL"
      .TextMatrix(0, 4) = "PESSOAL"
      .TextMatrix(0, 5) = "COMERCIAL"
      .TextMatrix(0, 6) = "COMERCIAL"
      .TextMatrix(0, 7) = "OP"
      .TextMatrix(0, 8) = "CELULAR"
      .TextMatrix(0, 9) = "OP"
      .TextMatrix(0, 10) = "CELULAR"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .ColAlignment(7) = 3
            .ColAlignment(9) = 3
            
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("nome")
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("residencial1"))
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("residencial2"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("comercial1"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("comercial2"))
            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("op1"))
            .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("celular1"))
            .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("op2"))
            .TextMatrix(.Rows - 1, 10) = ValidateNull(rTabela("celular2"))
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 7
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Command4_Click()
   Telefone_Cadastro.Show 1
End Sub

Private Sub Command6_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Atualizar_Grid
End Sub

Private Sub Form_Load()
   Atualizar_Grid
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   Dim telCad As Telefone_Cadastro
   
   Set telCad = New Telefone_Cadastro
   Load telCad
   telCad.cmdAlterar.Enabled = True
   telCad.cmdExcluir.Enabled = True
   telCad.txtCodigo.Text = ""
   telCad.txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
   'telCad.frmCadastro.Enabled = True
   'Telefone_Cadastro.txtNome.SetFocus
   
   telCad.Show vbModal
   
   Unload telCad
   Set telCad = Nothing
End Sub

Private Sub Image2_DblClick()
   Copyright.Show 1
End Sub

Private Sub txtNome_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM telefone WHERE (nome LIKE '%" & txtNome.Text & "%') ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   FormatarGrid r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
