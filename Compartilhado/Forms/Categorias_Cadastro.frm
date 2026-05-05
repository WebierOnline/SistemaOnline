VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Categorias_Cadastro 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Categorias"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8700
      TabIndex        =   2
      Top             =   0
      Width           =   8700
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00003399&
         BackStyle       =   0  'Transparent
         Caption         =   "CADASTRO DE CATEGORIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   8700
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridCategorias 
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   630
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   4763
      _Version        =   393216
      ScrollBars      =   2
   End
   Begin VB.Frame fraEntrada 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Dados da Categoria"
      Height          =   795
      Left            =   120
      TabIndex        =   10
      Top             =   3390
      Width           =   8460
      Begin VB.TextBox txtCategoria 
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   300
         Width           =   6735
      End
      Begin VB.Label lblCategoria 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lblAviso 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   690
         Width           =   8220
      End
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sair"
      Height          =   405
      Left            =   7380
      TabIndex        =   9
      Top             =   4260
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   405
      Left            =   5160
      TabIndex        =   8
      Top             =   4260
      Width           =   1200
   End
   Begin VB.CommandButton cmdExcluir 
      BackColor       =   &H000000FF&
      Caption         =   "Excluir"
      Height          =   405
      Left            =   3900
      TabIndex        =   7
      Top             =   4260
      Width           =   1200
   End
   Begin VB.CommandButton cmdEditar 
      BackColor       =   &H0000C0FF&
      Caption         =   "Editar"
      Height          =   405
      Left            =   2640
      TabIndex        =   6
      Top             =   4260
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalvar 
      BackColor       =   &H0000C000&
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1380
      TabIndex        =   5
      Top             =   4260
      Width           =   1200
   End
   Begin VB.CommandButton cmdNovo 
      BackColor       =   &H0000FF00&
      Caption         =   "Novo"
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   4260
      Width           =   1200
   End
   Begin VB.Label lblRegistros 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   4800
   End
End
Attribute VB_Name = "Categorias_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTipoEdicao As String
Dim vIDCategoria As Long
Dim tipoEmpresa As Integer

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Dim cCfg As Object
    Set cCfg = sysConfig("TIPO_EMPRESA")
    tipoEmpresa = CInt(cCfg.Value)
    DesabilitarEdicao
    ExibirGrid
End Sub

Private Sub ExibirGrid()
    Dim r As ADODB.Recordset
    RsOpen r, "SELECT ID_Categoria, Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria"
    FormatarGrid r
    If r.State <> 0 Then r.Close
    lblRegistros.Caption = gridCategorias.rows - 1 & " categoria(s) cadastrada(s)"
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
    With gridCategorias
        .Visible = False
        .Redraw = False
        .Clear
        .Cols = 2
        .rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .ColWidth(0) = 600
        .ColWidth(1) = 7740
        .TextMatrix(0, 0) = "Cód."
        .TextMatrix(0, 1) = "CATEGORIA"
        .Col = 0: .Row = 0: .CellFontBold = True: .CellAlignment = 4
        .Col = 1: .Row = 0: .CellFontBold = True: .CellAlignment = 4
        If Not rTabela Is Nothing Then
            Do While Not rTabela.EOF
                .TextMatrix(.rows - 1, 0) = rTabela("ID_Categoria")
                .TextMatrix(.rows - 1, 1) = rTabela("Categoria")
                rTabela.MoveNext
                .rows = .rows + 1
            Loop
        End If
        If .rows > 1 Then .rows = .rows - 1
        .Visible = True
        .Redraw = True
    End With
End Sub

Private Sub HabilitarEdicao()
    txtCategoria.Enabled = True
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNovo.Enabled = False
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    lblAviso.Caption = ""
End Sub

Private Sub DesabilitarEdicao()
    txtCategoria.Text = ""
    txtCategoria.Enabled = False
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNovo.Enabled = True
    cmdEditar.Enabled = True
    cmdExcluir.Enabled = True
    vIDCategoria = 0
    vTipoEdicao = ""
    lblAviso.Caption = ""
End Sub

Private Sub cmdNovo_Click()
    vTipoEdicao = "Novo"
    txtCategoria.Text = ""
    HabilitarEdicao
    txtCategoria.SetFocus
End Sub

Private Sub cmdEditar_Click()
    If gridCategorias.Row < 1 Then
        MsgBox "Selecione uma categoria para editar!", vbExclamation, "Atenção"
        Exit Sub
    End If
    vIDCategoria = CLng(gridCategorias.TextMatrix(gridCategorias.Row, 0))
    Dim r As ADODB.Recordset
    RsOpen r, "SELECT Categoria FROM Categorias WHERE ID_Categoria = " & vIDCategoria
    If Not r.EOF Then
        txtCategoria.Text = r("Categoria")
    End If
    If r.State <> 0 Then r.Close
    vTipoEdicao = "Edicao"
    HabilitarEdicao
    txtCategoria.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    If Trim(txtCategoria.Text) = "" Then
        lblAviso.Caption = "* Informe o nome da Categoria!"
        txtCategoria.SetFocus
        Exit Sub
    End If

    Dim vNome As String
    vNome = UCase(Trim(txtCategoria.Text))

    If vTipoEdicao = "Novo" Then
        SQLExecuta "INSERT INTO Categorias (Categoria, Tipo_Empresa) VALUES ('" & vNome & "', " & tipoEmpresa & ")"
        MsgBox "Categoria cadastrada com sucesso!", vbInformation, "Online Commerce"
    Else
        SQLExecuta "UPDATE Categorias SET Categoria = '" & vNome & "' WHERE ID_Categoria = " & vIDCategoria
        MsgBox "Categoria alterada com sucesso!", vbInformation, "Online Commerce"
    End If

    DesabilitarEdicao
    ExibirGrid
End Sub

Private Sub cmdExcluir_Click()
    If gridCategorias.Row < 1 Then
        MsgBox "Selecione uma categoria para excluir!", vbExclamation, "Atenção"
        Exit Sub
    End If
    vIDCategoria = CLng(gridCategorias.TextMatrix(gridCategorias.Row, 0))
    Dim vNomeExc As String
    vNomeExc = gridCategorias.TextMatrix(gridCategorias.Row, 1)
    If MsgBox("Excluir a categoria '" & vNomeExc & "'?" & vbCrLf & "Esta ação não pode ser desfeita.", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbNo Then Exit Sub
    SQLExecuta "DELETE FROM Categorias WHERE ID_Categoria = " & vIDCategoria
    MsgBox "Categoria excluída com sucesso!", vbInformation, "Online Commerce"
    ExibirGrid
End Sub

Private Sub cmdCancelar_Click()
    DesabilitarEdicao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub gridCategorias_DblClick()
    cmdEditar_Click
End Sub

Private Sub txtCategoria_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
