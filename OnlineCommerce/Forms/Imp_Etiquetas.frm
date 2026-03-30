VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Imp_Etiquentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SETORES"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12465
   Icon            =   "Imp_Etiquetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   7425
      TabIndex        =   4
      Top             =   60
      Width           =   7455
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Imp_Etiquetas.frx":23D2
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SETORES"
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
         Left            =   1140
         TabIndex        =   5
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.PictureBox frmSetor 
      Enabled         =   0   'False
      Height          =   795
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   5055
      TabIndex        =   1
      Top             =   960
      Width           =   5115
      Begin VB.ComboBox cboSetor 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   300
         Width           =   4995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   7680
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17648
            Text            =   "Desenv.: Online.Info Sistemas"
            TextSave        =   "Desenv.: Online.Info Sistemas"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "09:39"
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
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   615
      Left            =   11100
      TabIndex        =   7
      Top             =   1920
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Cancelar"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Imp_Etiquetas.frx":8742
      PICN            =   "Imp_Etiquetas.frx":875E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   615
      Left            =   11100
      TabIndex        =   8
      Top             =   2580
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Alterar"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Imp_Etiquetas.frx":A4F0
      PICN            =   "Imp_Etiquetas.frx":A50C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluir 
      Height          =   615
      Left            =   11100
      TabIndex        =   9
      Top             =   3300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Excluir"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Imp_Etiquetas.frx":C29E
      PICN            =   "Imp_Etiquetas.frx":C2BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdSalvar 
      Height          =   615
      Left            =   11100
      TabIndex        =   10
      Top             =   1260
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Salvar"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Imp_Etiquetas.frx":E04C
      PICN            =   "Imp_Etiquetas.frx":E068
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNovo 
      Height          =   615
      Left            =   11040
      TabIndex        =   11
      Top             =   600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Novo"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Imp_Etiquetas.frx":FDFA
      PICN            =   "Imp_Etiquetas.frx":FE16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5835
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   4234
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1215
      Left            =   7500
      TabIndex        =   13
      Top             =   4860
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2143
      _Version        =   393216
      TextStyleFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "Imp_Etiquentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private Function Inserir_Dados() As Boolean
Dim sSQL As String

sSQL = "INSERT INTO setor (cod_setor, setor) VALUES (" & _
   txtCodigo.Text & ", '" & cboSetor.Text & "');"

Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

sSQL = "UPDATE setor SET setor = '" & cboSetor.Text & "' WHERE (cod_setor = " & txtCodigo.Text & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(cod_setor), 0) AS codigo FROM setor;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodigo.Text = r("codigo") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer

With Grid
   .Clear
   .Cols = 3
   .Rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 2500
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "SETOR"
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = rTabela("cod_setor")
         .TextMatrix(.Rows - 1, 2) = rTabela("setor")
         rTabela.MoveNext
         
         .Rows = .Rows + 1
         i = i + 1
      Loop
   End If
   
   'MUDAR COR DE FONTE DA COLUNA
   'For i = 1 To .Rows - 1
   '   .Row = i
   '   .Col = 3
   '   .CellForeColor = &HC0&
   '   .CellFontBold = True
   'Next
   
   .Rows = .Rows - 1
   .Redraw = True
End With
End Sub

Private Sub Limpar_Objetos()
   txtCodigo.Text = ""
   cboSetor.Text = ""
End Sub

Private Sub cboSETOR_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Or cboSetor.Text = "" Then Exit Sub

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos
Form_Load
End Sub

Private Sub cmdExcluir_Click()
Dim sSQL As String
Dim bRet As Boolean

If txtCodigo.Text = "" Then Exit Sub

If ShowMsg("Excluir esse SETOR?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

sSQL = "DELETE FROM setor WHERE (cod_setor = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
Auto_Numeracao
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
frmSetor.Enabled = True
cboSetor.SetFocus
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo TrataErro

If txtCodigo.Text = "" Or cboSetor.Text = "" Then Exit Sub

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
MostrarGrid
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdNovo.Enabled = True
frmSetor.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
cboSetor.Text = (Grid.TextMatrix(Grid.Row, 2))
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmSetor.Enabled = True

End Sub

Private Sub MostrarGrid()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT cod_setor, setor FROM setor ORDER BY setor;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

