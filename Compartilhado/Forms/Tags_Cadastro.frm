VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Tags_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE TAGS"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFiltros 
      Caption         =   "Filtros"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9360
      Begin VB.TextBox txtFiltroTag 
         Height          =   315
         Left            =   660
         TabIndex        =   2
         Top             =   300
         Width           =   1800
      End
      Begin VB.ComboBox cboFiltroCategoria 
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Top             =   300
         Width           =   2400
      End
      Begin VB.CommandButton cmdLocalizar 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   270
         Width           =   1200
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Limpar"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label lblFiltroTag 
         Caption         =   "Tag:"
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblFiltroCategoria 
         Caption         =   "Categoria:"
         Height          =   252
         Left            =   2580
         TabIndex        =   3
         Top             =   330
         Width           =   840
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTags 
      Height          =   4200
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   7408
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.Frame frmDados 
      Caption         =   "Dados"
      Height          =   1440
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   9360
      Begin VB.TextBox txtTag 
         Height          =   315
         Left            =   660
         TabIndex        =   10
         Top             =   270
         Width           =   2400
      End
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Top             =   270
         Width           =   2400
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   870
         Width           =   840
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   870
         Width           =   840
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   870
         Width           =   840
      End
      Begin VB.Label lblTag 
         Caption         =   "Tag:"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categoria:"
         Height          =   252
         Left            =   3180
         TabIndex        =   11
         Top             =   300
         Width           =   840
      End
   End
End
Attribute VB_Name = "Tags_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Dim vTipo As String
Dim vIDTags As Long

Private Sub Form_Load()
   vTipo = "Novo"
   vIDTags = 0
   Set cCfg = sysConfig("TIPO_EMPRESA")
   tipoEmpresa = cCfg.Value
   Set cCfg = Nothing
   CarregarGrid
End Sub

Private Sub CarregarGrid()
Dim r As ADODB.Recordset
Dim sSQL As String
Dim sWhere As String
Dim lRow As Long

sWhere = " WHERE 1=1"
If Trim(txtFiltroTag.Text) <> "" Then
   sWhere = sWhere & " AND ct.Tags LIKE '%" & Replace(Trim(txtFiltroTag.Text), "'", "''") & "%'"
End If
If Trim(cboFiltroCategoria.Text) <> "" Then
   sWhere = sWhere & " AND c.Categoria = '" & Replace(Trim(cboFiltroCategoria.Text), "'", "''") & "'"
End If

sSQL = "SELECT ct.ID_Tags, ct.Tags, c.Categoria FROM Categorias_Tags ct " & _
       "INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria" & _
       sWhere & " ORDER BY c.Categoria, ct.Tags"

With grdTags
   .Redraw = False
   .Clear
   .Cols = 3
   .rows = 2
   .ColWidth(0) = 0
   .ColWidth(1) = 4200
   .ColWidth(2) = 4200
   .FixedRows = 1
   .FixedCols = 0
   .TextMatrix(0, 1) = "TAG"
   .TextMatrix(0, 2) = "CATEGORIA"
   .Row = 0
   .Col = 0
   .ColSel = .Cols - 1
   .CellFontBold = True
   .FillStyle = 1

   Set r = dbData.OpenRecordset(sSQL)
   Do While Not r.EOF
      .TextMatrix(.rows - 1, 0) = ValidateNull(r("ID_Tags"))
      .TextMatrix(.rows - 1, 1) = ValidateNull(r("Tags"))
      .TextMatrix(.rows - 1, 2) = ValidateNull(r("Categoria"))
      r.MoveNext
      .rows = .rows + 1
   Loop
   .rows = .rows - 1
   If r.State <> 0 Then r.Close

   For lRow = 1 To .rows - 1
      .Row = lRow
      .Col = 0
      .ColSel = .Cols - 1
      If lRow Mod 2 = 0 Then
         .CellBackColor = &HE0E0E0
      Else
         .CellBackColor = vbWhite
      End If
   Next lRow
   .FillStyle = 0
   .Redraw = True
End With
End Sub

Private Sub cboFiltroCategoria_GotFocus()
Dim vAnt As String
vAnt = cboFiltroCategoria.Text
cboFiltroCategoria.Clear
Dim r As ADODB.Recordset
Set r = dbData.OpenRecordset("SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria")
Do While Not r.EOF
   cboFiltroCategoria.AddItem ValidateNull(r("Categoria"))
   r.MoveNext
Loop
If r.State <> 0 Then r.Close
cboFiltroCategoria.Text = vAnt
End Sub

Private Sub cmdLocalizar_Click()
   CarregarGrid
End Sub

Private Sub cmdLimpar_Click()
   txtFiltroTag.Text = ""
   cboFiltroCategoria.Text = ""
   CarregarGrid
End Sub

Private Sub grdTags_Click()
   If grdTags.Row < 1 Then Exit Sub
   vIDTags = CLng(grdTags.TextMatrix(grdTags.Row, 0))
   txtTag.Text = grdTags.TextMatrix(grdTags.Row, 1)
   cboCategoria.Text = grdTags.TextMatrix(grdTags.Row, 2)
   vTipo = "Edicao"
End Sub

Private Sub cboCategoria_GotFocus()
Dim vAnt As String
vAnt = cboCategoria.Text
cboCategoria.Clear
Dim r As ADODB.Recordset
Set r = dbData.OpenRecordset("SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria")
Do While Not r.EOF
   cboCategoria.AddItem ValidateNull(r("Categoria"))
   r.MoveNext
Loop
If r.State <> 0 Then r.Close
cboCategoria.Text = vAnt
End Sub

Private Sub cmdNovo_Click()
   vTipo = "Novo"
   vIDTags = 0
   txtTag.Text = ""
   cboCategoria.Text = ""
   txtTag.SetFocus
End Sub

Private Sub cmdSalvar_Click()
Dim r As ADODB.Recordset
Dim lCatID As Long

If Trim(txtTag.Text) = "" Then
   MsgBox "Digite o nome da Tag.", vbInformation, "Aviso"
   txtTag.SetFocus
   Exit Sub
End If
If Trim(cboCategoria.Text) = "" Then
   MsgBox "Selecione a Categoria.", vbInformation, "Aviso"
   cboCategoria.SetFocus
   Exit Sub
End If

Set r = dbData.OpenRecordset("SELECT ID_Categoria FROM Categorias WHERE Categoria = '" & Replace(cboCategoria.Text, "'", "''") & "'")
If Not r.EOF Then lCatID = CLng(r("ID_Categoria"))
If r.State <> 0 Then r.Close

If lCatID = 0 Then
   MsgBox "Categoria não encontrada.", vbInformation, "Aviso"
   Exit Sub
End If

If vTipo = "Novo" Then
   dbData.Execute "INSERT INTO Categorias_Tags (Tags, ID_Categoria) VALUES ('" & Replace(Trim(txtTag.Text), "'", "''") & "', " & lCatID & ");"
Else
   dbData.Execute "UPDATE Categorias_Tags SET Tags = '" & Replace(Trim(txtTag.Text), "'", "''") & "', ID_Categoria = " & lCatID & " WHERE ID_Tags = " & vIDTags & ";"
End If

cmdNovo_Click
CarregarGrid
End Sub

Private Sub cmdExcluir_Click()
   If vIDTags = 0 Then Exit Sub
   If MsgBox("Excluir a tag '" & txtTag.Text & "'?", vbQuestion + vbYesNo, "Confirmar") = vbNo Then Exit Sub
   dbData.Execute "DELETE FROM Categorias_Tags WHERE ID_Tags = " & vIDTags & ";"
   cmdNovo_Click
   CarregarGrid
End Sub
