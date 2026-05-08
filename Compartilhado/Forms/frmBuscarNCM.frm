VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBuscarNCM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar NCM"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   450
      Left            =   9720
      TabIndex        =   11
      Top             =   6480
      Width           =   2400
   End
   Begin VB.CommandButton cmdUsarNCM 
      Caption         =   "Usar NCM"
      Height          =   450
      Left            =   60
      TabIndex        =   10
      Top             =   6480
      Width           =   2400
   End
   Begin MSFlexGridLib.MSFlexGrid lstProdutos 
      Height          =   5340
      Left            =   60
      TabIndex        =   9
      Top             =   1080
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   9419
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame fraFiltro 
      Height          =   585
      Left            =   60
      TabIndex        =   12
      Top             =   480
      Width           =   12060
      Begin VB.TextBox txtNCMF 
         Height          =   315
         Left            =   9900
         TabIndex        =   8
         Top             =   180
         Width           =   1140
      End
      Begin VB.OptionButton optNCM 
         Caption         =   "NCM:"
         Height          =   195
         Left            =   9120
         TabIndex        =   7
         Top             =   180
         Width           =   720
      End
      Begin VB.TextBox txtDescF 
         Height          =   315
         Left            =   4860
         TabIndex        =   6
         Top             =   180
         Width           =   3960
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   3720
         TabIndex        =   5
         Top             =   180
         Width           =   1080
      End
      Begin VB.TextBox txtCodBarraF 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   180
         Width           =   2280
      End
      Begin VB.OptionButton optCodBarra 
         Caption         =   "Cód. Barra:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   9420
      TabIndex        =   2
      Top             =   60
      Width           =   1380
   End
   Begin VB.ComboBox cboTagsF 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2940
   End
   Begin VB.ComboBox cboCategoriaF 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4200
   End
   Begin VB.Label lblTagsF 
      AutoSize        =   -1  'True
      Caption         =   "Tags:"
      Height          =   195
      Left            =   5760
      TabIndex        =   14
      Top             =   126
      Width           =   480
   End
   Begin VB.Label lblCategoriaF 
      AutoSize        =   -1  'True
      Caption         =   "Categoria:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   126
      Width           =   840
   End
   Begin VB.Frame fraCarregando 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ""
      Height          =   600
      Left            =   3090
      TabIndex        =   50
      Top             =   3200
      Visible         =   0   'False
      Width           =   4800
      Begin VB.Label lblCarregando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   0
         Top             =   90
         Width           =   4680
      End
   End
   Begin VB.Timer tmrCarregando 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   120
      Top             =   6840
   End
End
Attribute VB_Name = "frmBuscarNCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sNCMSelecionado As String
Public sCatInicial As String
Public sTagInicial As String
Private iDots As Integer

Private Sub Form_Load()
    Dim i As Integer
    sNCMSelecionado = ""
    iDots = 0
    fraCarregando.Visible = False
    'optDesc.Value = True
    CarregarCategorias
    If sCatInicial <> "" Then
        For i = 0 To cboCategoriaF.ListCount - 1
            If cboCategoriaF.List(i) = sCatInicial Then
                cboCategoriaF.ListIndex = i: Exit For
            End If
        Next i
        CarregarTags
        If sTagInicial <> "" Then
            For i = 0 To cboTagsF.ListCount - 1
                If cboTagsF.List(i) = sTagInicial Then
                    cboTagsF.ListIndex = i: Exit For
                End If
            Next i
        End If
    End If
    CarregarGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        sNCMSelecionado = ""
        Me.Hide
    End If
End Sub

' --- Opcoes de pesquisa ---
Private Sub optCodBarra_Click()
    txtCodBarraF.SetFocus
    txtDescF.Text = ""
    txtNCMF.Text = ""
End Sub

Private Sub optDesc_Click()
    txtDescF.SetFocus
    txtCodBarraF.Text = ""
    txtNCMF.Text = ""
End Sub

Private Sub optNCM_Click()
    txtNCMF.SetFocus
    txtCodBarraF.Text = ""
    txtDescF.Text = ""
End Sub

' --- Carga de dados ---
Private Sub CarregarCategorias()
    Dim rCat As ADODB.Recordset
    cboCategoriaF.Clear
    cboCategoriaF.AddItem ""
    RsOpen rCat, "SELECT DISTINCT Categoria FROM Categorias ORDER BY Categoria"
    Do While Not rCat.EOF
        cboCategoriaF.AddItem Trim(rCat("Categoria"))
        rCat.MoveNext
    Loop
    If rCat.State <> 0 Then rCat.Close
End Sub

Private Sub CarregarTags()
    Dim rTags As ADODB.Recordset
    cboTagsF.Clear
    cboTagsF.AddItem "TODOS"
    If cboCategoriaF.Text = "" Then Exit Sub
    RsOpen rTags, "SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> '' AND categoria = '" & Replace(cboCategoriaF.Text, "'", "''") & "' ORDER BY TAGS"
    Do While Not rTags.EOF
        cboTagsF.AddItem Trim(rTags("TAGS"))
        rTags.MoveNext
    Loop
    If rTags.State <> 0 Then rTags.Close
End Sub

Private Sub ConfigurarGrid()
    With lstProdutos
        .Cols = 6
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(2) = 5600
        .ColWidth(3) = 1000
        .ColWidth(4) = 2100
        .ColWidth(5) = 1500
        .Row = 0: .Col = 0: .Text = "codigo"
        .Row = 0: .Col = 1: .Text = "Cod. Barra"
        .Row = 0: .Col = 2: .Text = "Descrição"
        .Row = 0: .Col = 3: .Text = "NCM": .CellBackColor = &HC0C0C0
        .Row = 0: .Col = 4: .Text = "Categoria"
        .Row = 0: .Col = 5: .Text = "Tags"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
End Sub

Private Sub CarregarGrid()
    Dim rProd As ADODB.Recordset
    Dim sql As String, w As String
    Dim r As Integer
    Dim bTextSearch As Boolean

    fraCarregando.Visible = True
    fraCarregando.ZOrder 0
    tmrCarregando.Enabled = True
    Me.Refresh
    DoEvents
    ConfigurarGrid
    lstProdutos.rows = 1

    bTextSearch = False
    w = "WHERE 1=1"

    ' Campo de texto ativo desvincula filtros de combo
    If optCodBarra.Value And Trim(txtCodBarraF.Text) <> "" Then
        w = w & " AND cod_barra LIKE '%" & Replace(Trim(txtCodBarraF.Text), "'", "''") & "%'"
        bTextSearch = True
    ElseIf optDesc.Value And Trim(txtDescF.Text) <> "" Then
        w = w & " AND descricao LIKE '%" & Replace(Trim(txtDescF.Text), "'", "''") & "%'"
        bTextSearch = True
    ElseIf optNCM.Value And Trim(txtNCMF.Text) <> "" Then
        w = w & " AND NCM LIKE '%" & Replace(Trim(txtNCMF.Text), "'", "''") & "%'"
        bTextSearch = True
    End If

    If Not bTextSearch Then
        If cboCategoriaF.Text <> "" Then
            w = w & " AND categoria = '" & Replace(cboCategoriaF.Text, "'", "''") & "'"
        End If
        If cboTagsF.Text <> "" And cboTagsF.Text <> "TODOS" Then
            w = w & " AND TAGS = '" & Replace(cboTagsF.Text, "'", "''") & "'"
        End If
    End If

    sql = "SELECT TOP 300 codigo, cod_barra, descricao, NCM, categoria, ISNULL(TAGS,'') AS TAGS " & _
          "FROM produtos " & w & " ORDER BY descricao"
    RsOpen rProd, sql
    r = 1
    Do While Not rProd.EOF
        With lstProdutos
            .rows = r + 1
            .Row = r: .Col = 0: .Text = CStr(rProd("codigo"))
            .Row = r: .Col = 1: .Text = Trim(IIf(IsNull(rProd("cod_barra")), "", rProd("cod_barra")))
            .Row = r: .Col = 2: .Text = Trim(IIf(IsNull(rProd("descricao")), "", rProd("descricao")))
            .Row = r: .Col = 3: .Text = Trim(IIf(IsNull(rProd("NCM")), "", rProd("NCM"))): .CellBackColor = &HE0E0E0
            .Row = r: .Col = 4: .Text = Trim(IIf(IsNull(rProd("categoria")), "", rProd("categoria")))
            .Row = r: .Col = 5: .Text = Trim(rProd("TAGS"))
        End With
        r = r + 1
        rProd.MoveNext
    Loop
    If rProd.State <> 0 Then rProd.Close
    tmrCarregando.Enabled = False
    fraCarregando.Visible = False
End Sub

' --- Eventos dos combos ---
Private Sub cboCategoriaF_Click()
    CarregarTags
    CarregarGrid
End Sub

Private Sub cboTagsF_Click()
    CarregarGrid
End Sub

Private Sub cmdFiltrar_Click()
    CarregarGrid
End Sub

' --- Selecao e captura do NCM ---
Private Sub lstProdutos_DblClick()
    cmdUsarNCM_Click
End Sub

Private Sub cmdUsarNCM_Click()
    If lstProdutos.Row < 1 Then MsgBox "Selecione um produto.", vbInformation: Exit Sub
    Dim sNCM As String
    sNCM = Trim(lstProdutos.TextMatrix(lstProdutos.Row, 3))
    If sNCM = "" Then MsgBox "Produto sem NCM.", vbInformation: Exit Sub
    sNCMSelecionado = sNCM
    Me.Hide
End Sub

Private Sub cmdFechar_Click()
    sNCMSelecionado = ""
    Me.Hide
End Sub

Private Sub tmrCarregando_Timer()
    iDots = (iDots Mod 3) + 1
    lblCarregando.Caption = "Carregando" & String(iDots, ".")
End Sub
