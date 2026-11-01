VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Tags_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE TAGS"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFiltros 
      Caption         =   "Filtros"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   5520
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
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   375
         Left            =   5940
         TabIndex        =   16
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Exibir"
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
         MICON           =   "Tags_Cadastro.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdLimpar 
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Limpar"
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
         MICON           =   "Tags_Cadastro.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   300
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTags 
      Height          =   4200
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   7408
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.Frame frmDados 
      Caption         =   "Dados"
      Height          =   1200
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   9360
      Begin VB.TextBox txtTag 
         Height          =   315
         Left            =   540
         TabIndex        =   8
         Top             =   270
         Width           =   2640
      End
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   4140
         TabIndex        =   10
         Top             =   270
         Width           =   3300
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   375
         Left            =   540
         TabIndex        =   11
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Novo"
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
         MICON           =   "Tags_Cadastro.frx":0038
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
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Salvar"
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
         MICON           =   "Tags_Cadastro.frx":0054
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
         Height          =   375
         Left            =   5100
         TabIndex        =   13
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Excluir"
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
         MICON           =   "Tags_Cadastro.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   375
         Left            =   2820
         TabIndex        =   14
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "Tags_Cadastro.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdEditar 
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Editar"
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
         MICON           =   "Tags_Cadastro.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblTag 
         Caption         =   "Tag:"
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categoria:"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
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
Dim bEditando As Boolean   'True = modo Novo ou Edicao (cmdSalvar/cmdCancelar ativos)
Dim vLinhaSelecionada As Long   '0 = nenhuma linha do grid selecionada

Private Sub Form_Load()
   Call ResetarParaOcioso
   Set cCfg = sysConfig("TIPO_EMPRESA")
   tipoEmpresa = cCfg.Value
   Set cCfg = Nothing
   CarregarGrid
   Call AtualizarBotoes
End Sub

Private Sub ResetarParaOcioso()
'Estado parado: nada sendo editado, nenhuma linha destacada.
bEditando = False
vTipo = "Novo"
vIDTags = 0
vLinhaSelecionada = 0
txtTag.Text = ""
cboCategoria.Text = ""
End Sub

Private Sub AtualizarBotoes()
'cmdSalvar/cmdCancelar so ativos durante edicao (Novo ou Edicao); os
'demais so ativos fora da edicao, e cmdEditar/cmdExcluir exigem uma
'linha selecionada no grid.
Dim bLinhaSelecionada As Boolean
bLinhaSelecionada = (vIDTags <> 0)

If bEditando Then
    cmdNovo.Enabled = False
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
Else
    cmdNovo.Enabled = True
    cmdEditar.Enabled = bLinhaSelecionada
    cmdExcluir.Enabled = bLinhaSelecionada
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
End If
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
'bloqueado durante edicao - trocar de linha selecionada no meio de um
'Novo/Edicao sobrescreveria vIDTags e o Salvar atualizaria o registro errado.
Dim lNovaLinha As Long

If bEditando Then Exit Sub
If grdTags.Row < 1 Then Exit Sub

'captura a linha clicada ANTES de mexer em .Row pra qualquer coisa - .Row
'e a mesma propriedade usada tanto pra navegar quanto pra pintar, entao
'reaproveitar sem capturar antes sobrescreve a linha recem-clicada.
lNovaLinha = grdTags.Row

'restaura a cor zebra da linha que estava destacada antes (se houver e for diferente)
If vLinhaSelecionada >= 1 And vLinhaSelecionada <> lNovaLinha Then
   With grdTags
      .Row = vLinhaSelecionada
      .Col = 0
      .ColSel = .Cols - 1
      If vLinhaSelecionada Mod 2 = 0 Then
         .CellBackColor = &HE0E0E0
      Else
         .CellBackColor = vbWhite
      End If
      .CellForeColor = vbBlack
   End With
End If

vLinhaSelecionada = lNovaLinha
vIDTags = CLng(grdTags.TextMatrix(lNovaLinha, 0))

With grdTags
   .Row = lNovaLinha
   .Col = 0
   .ColSel = .Cols - 1
   .CellBackColor = vbBlue
   .CellForeColor = vbWhite
End With

Call AtualizarBotoes
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
   bEditando = True
   Call AtualizarBotoes
   txtTag.SetFocus
End Sub

Private Sub cmdEditar_Click()
If vLinhaSelecionada < 1 Then Exit Sub
vTipo = "Edicao"
txtTag.Text = grdTags.TextMatrix(vLinhaSelecionada, 1)
cboCategoria.Text = grdTags.TextMatrix(vLinhaSelecionada, 2)
bEditando = True
Call AtualizarBotoes
txtTag.SetFocus
End Sub

Private Sub cmdCancelar_Click()
'desiste do cadastro/edicao em andamento, sem salvar nada
Call ResetarParaOcioso
CarregarGrid
Call AtualizarBotoes
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

Call ResetarParaOcioso
CarregarGrid
Call AtualizarBotoes
End Sub

Private Sub cmdExcluir_Click()
   If vIDTags = 0 Then Exit Sub
   If MsgBox("Excluir a tag '" & grdTags.TextMatrix(vLinhaSelecionada, 1) & "'?", vbQuestion + vbYesNo, "Confirmar") = vbNo Then Exit Sub
   dbData.Execute "DELETE FROM Categorias_Tags WHERE ID_Tags = " & vIDTags & ";"
   Call ResetarParaOcioso
   CarregarGrid
   Call AtualizarBotoes
End Sub

Private Sub txtFiltroTag_Change()

End Sub

Private Sub txtFiltroTag_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtTag_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


