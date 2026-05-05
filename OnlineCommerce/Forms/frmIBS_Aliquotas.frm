VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmIBS_Aliquotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IBS - Alíquotas por Estado e Município"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEstado 
      Caption         =   "Alíquota IBS por Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   13260
      Begin VB.CommandButton cmdAplicarTodosUF 
         Caption         =   "Aplicar para TODOS os Estados"
         Height          =   360
         Left            =   10260
         TabIndex        =   45
         Top             =   2880
         Width           =   2880
      End
      Begin VB.CommandButton cmdExcluirEstado 
         Caption         =   "Excluir"
         Height          =   360
         Left            =   11700
         TabIndex        =   44
         Top             =   1560
         Width           =   1440
      End
      Begin VB.CommandButton cmdSalvarEstado 
         Caption         =   "Salvar"
         Height          =   360
         Left            =   10200
         TabIndex        =   43
         Top             =   1560
         Width           =   1440
      End
      Begin VB.CommandButton cmdNovoEstado 
         Caption         =   "Novo"
         Height          =   360
         Left            =   8700
         TabIndex        =   42
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Frame Frame2 
         Caption         =   "Edição Individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   8280
         TabIndex        =   33
         Top             =   240
         Width           =   4875
         Begin VB.TextBox txtDFimUFEdit 
            Height          =   288
            Left            =   3480
            TabIndex        =   41
            Top             =   660
            Width           =   1020
         End
         Begin VB.TextBox txtDIniUFEdit 
            Height          =   288
            Left            =   1980
            TabIndex        =   39
            Top             =   660
            Width           =   1020
         End
         Begin VB.TextBox txtAliqUFEdit 
            Height          =   288
            Left            =   720
            TabIndex        =   37
            Top             =   660
            Width           =   660
         End
         Begin VB.ComboBox cboUFEdit 
            Height          =   315
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   300
            Width           =   765
         End
         Begin VB.Label lblDFimUFEdit 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
            Height          =   195
            Left            =   3120
            TabIndex        =   40
            Top             =   660
            Width           =   285
         End
         Begin VB.Label lblDIniUFEdit 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   1500
            TabIndex        =   38
            Top             =   660
            Width           =   420
         End
         Begin VB.Label lblAliqUFEdit 
            AutoSize        =   -1  'True
            Caption         =   "Alíq. %:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblUFEditE 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   60
         TabIndex        =   28
         Top             =   240
         Width           =   8175
         Begin VB.CommandButton cmdTodosEstado 
            Caption         =   "Todos"
            Height          =   288
            Left            =   3720
            TabIndex        =   32
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdFiltrarEstado 
            Caption         =   "Filtrar"
            Height          =   288
            Left            =   2460
            TabIndex        =   31
            Top             =   240
            Width           =   1200
         End
         Begin VB.ComboBox cboFiltroUF 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   240
            Width           =   1500
         End
         Begin MSFlexGridLib.MSFlexGrid lstEstado 
            Height          =   3180
            Left            =   60
            TabIndex        =   46
            Top             =   600
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   5609
            _Version        =   393216
         End
         Begin VB.Label lblFiltroUF 
            AutoSize        =   -1  'True
            Caption         =   "Filtrar UF:"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   300
            Width           =   675
         End
      End
      Begin VB.Frame fraAplicarTodosUF 
         Caption         =   "Edição de TODOS os Estados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   8340
         TabIndex        =   1
         Top             =   2040
         Width           =   4800
         Begin VB.TextBox txtAliqTodosUF 
            Height          =   288
            Left            =   720
            TabIndex        =   3
            Top             =   300
            Width           =   660
         End
         Begin VB.TextBox txtDIniTodosUF 
            Height          =   288
            Left            =   2040
            TabIndex        =   5
            Top             =   300
            Width           =   1020
         End
         Begin VB.TextBox txtDFimTodosUF 
            Height          =   288
            Left            =   3660
            TabIndex        =   7
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label lblAliqTodosUF 
            AutoSize        =   -1  'True
            Caption         =   "Alíq. %:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblDIniTodosUF 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   1560
            TabIndex        =   4
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblDFimTodosUF 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
            Height          =   195
            Left            =   3300
            TabIndex        =   6
            Top             =   300
            Width           =   285
         End
      End
      Begin VB.Label lblCarregando 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CARREGANDO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   8280
         TabIndex        =   47
         Top             =   3480
         Visible         =   0   'False
         Width           =   4800
      End
   End
   Begin VB.CommandButton cmdAtualizarCidade 
      Caption         =   "Atualizar Cidade com Alíquotas Vigentes de Hoje"
      Height          =   480
      Left            =   120
      TabIndex        =   25
      Top             =   12600
      Width           =   5760
   End
   Begin VB.Frame fraMunicipio 
      Caption         =   "Alíquota IBS por Município"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   60
      TabIndex        =   8
      Top             =   4320
      Width           =   13260
      Begin VB.CommandButton cmdAplicarTodosMun 
         Caption         =   "Aplicar"
         Height          =   360
         Left            =   11700
         TabIndex        =   62
         Top             =   4920
         Width           =   1440
      End
      Begin VB.CommandButton cmdExcluirMun 
         Caption         =   "Excluir"
         Height          =   360
         Left            =   5160
         TabIndex        =   61
         Top             =   4920
         Width           =   1440
      End
      Begin VB.CommandButton cmdSalvarMun 
         Caption         =   "Salvar"
         Height          =   360
         Left            =   3660
         TabIndex        =   60
         Top             =   4920
         Width           =   1440
      End
      Begin VB.CommandButton cmdNovoMun 
         Caption         =   "Novo"
         Height          =   360
         Left            =   2160
         TabIndex        =   59
         Top             =   4920
         Width           =   1440
      End
      Begin VB.Frame Frame3 
         Caption         =   "Edição Individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   60
         TabIndex        =   48
         Top             =   3840
         Width           =   6555
         Begin VB.TextBox txtDFimMunEdit 
            Height          =   288
            Left            =   3660
            TabIndex        =   58
            Top             =   600
            Width           =   1020
         End
         Begin VB.TextBox txtDIniMunEdit 
            Height          =   288
            Left            =   2160
            TabIndex        =   56
            Top             =   600
            Width           =   1020
         End
         Begin VB.TextBox txtAliqMunEdit 
            Height          =   288
            Left            =   900
            TabIndex        =   54
            Top             =   600
            Width           =   660
         End
         Begin VB.TextBox txtNomeMunEdit 
            Height          =   288
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   300
            Width           =   3900
         End
         Begin VB.TextBox txtCodMunEdit 
            Height          =   288
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lblDFimMunEdit 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
            Height          =   195
            Left            =   3300
            TabIndex        =   57
            Top             =   600
            Width           =   285
         End
         Begin VB.Label lblDIniMunEdit 
            Caption         =   "Inicio:"
            Height          =   240
            Left            =   1680
            TabIndex        =   55
            Top             =   600
            Width           =   840
         End
         Begin VB.Label lblAliqMunEdit 
            AutoSize        =   -1  'True
            Caption         =   "Alíq. %:"
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblNomeMunEdit 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   1980
            TabIndex        =   51
            Top             =   300
            Width           =   465
         End
         Begin VB.Label lblCodMunEdit 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Mun:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.ComboBox cboFiltroUFMun 
         Height          =   288
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtFiltroCidade 
         Height          =   288
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   2880
      End
      Begin VB.CommandButton cmdFiltrarMun 
         Caption         =   "Filtrar"
         Height          =   288
         Left            =   5880
         TabIndex        =   13
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton cmdLimparMun 
         Caption         =   "Limpar"
         Height          =   288
         Left            =   7200
         TabIndex        =   14
         Top             =   360
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid lstMunicipio 
         Height          =   3060
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   13080
         _ExtentX        =   23072
         _ExtentY        =   5398
         _Version        =   393216
      End
      Begin VB.Frame fraAplicarTodosMun 
         Caption         =   "Aplicar Alíquota em Massa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   6660
         TabIndex        =   16
         Top             =   3840
         Width           =   6480
         Begin VB.OptionButton optTodosMun 
            Caption         =   "Todos os municípios"
            Height          =   240
            Left            =   2220
            TabIndex        =   17
            Top             =   240
            Width           =   1920
         End
         Begin VB.OptionButton optEstadoFiltrado 
            Caption         =   "Somente estado filtrado"
            Height          =   240
            Left            =   60
            TabIndex        =   18
            Top             =   240
            Width           =   2100
         End
         Begin VB.TextBox txtAliqTodosMun 
            Height          =   288
            Left            =   720
            TabIndex        =   20
            Top             =   540
            Width           =   660
         End
         Begin VB.TextBox txtDIniTodosMun 
            Height          =   288
            Left            =   1980
            TabIndex        =   22
            Top             =   540
            Width           =   1020
         End
         Begin VB.TextBox txtDFimTodosMun 
            Height          =   288
            Left            =   3480
            TabIndex        =   24
            Top             =   540
            Width           =   1020
         End
         Begin VB.Label lblAliqTodosMun 
            Caption         =   "Alíq. %:"
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lblDIniTodosMun 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   1500
            TabIndex        =   21
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblDFimTodosMun 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
            Height          =   195
            Left            =   3120
            TabIndex        =   23
            Top             =   540
            Width           =   285
         End
      End
      Begin VB.Label lblFiltroUFMun 
         Caption         =   "UF:"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblFiltroCidade 
         Caption         =   "Cidade:"
         Height          =   240
         Left            =   2100
         TabIndex        =   11
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label lblStatusSync 
      Height          =   480
      Left            =   6120
      TabIndex        =   26
      Top             =   12600
      Width           =   8160
   End
   Begin VB.Label lblAvisoSync 
      Height          =   240
      Left            =   120
      TabIndex        =   27
      Top             =   12360
      Width           =   14160
   End
End
Attribute VB_Name = "frmIBS_Aliquotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bNovoEstado As Boolean
Dim bNovoMun    As Boolean
Dim iEstadoId   As Long
Dim iMunId      As Long
Dim sCodMunSel  As String

' ============================================================
' FORM LOAD
' ============================================================
Private Sub Form_Load()
    CarregarComboUF cboFiltroUF, True
    CarregarComboUF cboUFEdit, False
    CarregarComboUF cboFiltroUFMun, True
    optTodosMun.Value = True
    ' Carrega municipios so do primeiro estado por padrao
    If cboFiltroUFMun.ListCount > 1 Then cboFiltroUFMun.ListIndex = 1

    With lstEstado
        .Cols = 5
        .ColWidth(0) = 840:  .Row = 0: .Col = 0: .Text = "UF"
        .ColWidth(1) = 1440: .Row = 0: .Col = 1: .Text = "Al" & Chr(237) & "q. %"
        .ColWidth(2) = 1680: .Row = 0: .Col = 2: .Text = "Vig. Ini"
        .ColWidth(3) = 1680: .Row = 0: .Col = 3: .Text = "Vig. Fim"
        .ColWidth(4) = 0     ' Id oculto
        .AllowUserResizing = 1: .SelectionMode = 1
    End With

    With lstMunicipio
        .Cols = 7
        .ColWidth(0) = 720:  .Row = 0: .Col = 0: .Text = "UF"
        .ColWidth(1) = 960:  .Row = 0: .Col = 1: .Text = "Cod. Mun"
        .ColWidth(2) = 5040: .Row = 0: .Col = 2: .Text = "Nome"
        .ColWidth(3) = 1200: .Row = 0: .Col = 3: .Text = "Al" & Chr(237) & "q. %"
        .ColWidth(4) = 1560: .Row = 0: .Col = 4: .Text = "Vig. Ini"
        .ColWidth(5) = 1560: .Row = 0: .Col = 5: .Text = "Vig. Fim"
        .ColWidth(6) = 0     ' Id oculto
        .AllowUserResizing = 1: .SelectionMode = 1
    End With

    CarregarIBSEstado
    CarregarIBSMunicipio
    VerificarAliquotasPendentes
End Sub

Private Sub CarregarComboUF(cbo As ComboBox, comTodos As Boolean)
    Dim rRs As ADODB.Recordset
    cbo.Clear
    If comTodos Then cbo.AddItem "(Todos)"
    RsOpen rRs, "SELECT DISTINCT UF FROM Cidade WHERE UF IS NOT NULL ORDER BY UF"
    Do While Not rRs.EOF
        cbo.AddItem rRs("UF")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    cbo.ListIndex = 0
End Sub

' ============================================================
' IBS ESTADO
' ============================================================
Private Sub CarregarIBSEstado()
    Dim rRs As ADODB.Recordset
    Dim sql As String, r As Long
    sql = "SELECT Id, UF, IBSUFpAliq, dIniVig, dFimVig FROM IBS_Estado"
    If cboFiltroUF.ListIndex > 0 Then
        sql = sql & " WHERE UF = '" & cboFiltroUF.Text & "'"
    End If
    sql = sql & " ORDER BY UF, dIniVig"
    lstEstado.rows = 1
    RsOpen rRs, sql
    Do While Not rRs.EOF
        lstEstado.AddItem ""
        r = lstEstado.rows - 1
        lstEstado.Row = r: lstEstado.Col = 0: lstEstado.Text = rRs("UF")
        lstEstado.Row = r: lstEstado.Col = 1: lstEstado.Text = Format(rRs("IBSUFpAliq"), "0.00")
        lstEstado.Row = r: lstEstado.Col = 2: lstEstado.Text = IIf(IsNull(rRs("dIniVig")), "", Format(rRs("dIniVig"), "dd/mm/yyyy"))
        lstEstado.Row = r: lstEstado.Col = 3: lstEstado.Text = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "dd/mm/yyyy"))
        lstEstado.Row = r: lstEstado.Col = 4: lstEstado.Text = rRs("Id")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub lstEstado_Click()
    If lstEstado.Row = 0 Then Exit Sub
    iEstadoId = Val(lstEstado.TextMatrix(lstEstado.Row, 4))
    cboUFEdit.Text = lstEstado.TextMatrix(lstEstado.Row, 0)
    txtAliqUFEdit.Text = lstEstado.TextMatrix(lstEstado.Row, 1)
    txtDIniUFEdit.Text = lstEstado.TextMatrix(lstEstado.Row, 2)
    txtDFimUFEdit.Text = lstEstado.TextMatrix(lstEstado.Row, 3)
    bNovoEstado = False
End Sub

Private Sub cmdFiltrarEstado_Click(): CarregarIBSEstado: End Sub
Private Sub cmdTodosEstado_Click(): cboFiltroUF.ListIndex = 0: CarregarIBSEstado: End Sub

Private Sub cmdNovoEstado_Click()
    bNovoEstado = True: iEstadoId = 0
    cboUFEdit.ListIndex = 0: txtAliqUFEdit.Text = "": txtDIniUFEdit.Text = "": txtDFimUFEdit.Text = ""
    cboUFEdit.SetFocus
End Sub

Private Sub cmdSalvarEstado_Click()
    Dim sUF As String, sAliq As String, sDIni As String, sDFim As String
    sUF = Trim(cboUFEdit.Text)
    sAliq = Replace(Trim(txtAliqUFEdit.Text), ",", ".")
    sDIni = Trim(txtDIniUFEdit.Text)
    sDFim = Trim(txtDFimUFEdit.Text)
    If sUF = "" Or sAliq = "" Or sDIni = "" Then
        MsgBox "UF, Al" & Chr(237) & "q e Vig. Ini s" & Chr(227) & "o obrigat" & Chr(243) & "rios.", vbExclamation: Exit Sub
    End If
    Dim sIdEstado As String
    sIdEstado = SQLExecutaRetorno("SELECT TOP 1 CAST(IdEstado AS VARCHAR) AS IdEstado FROM Cidade WHERE UF='" & sUF & "'", "IdEstado", "")
    If sIdEstado = "" Then MsgBox "UF n" & Chr(227) & "o encontrada em Cidade.", vbExclamation: Exit Sub
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    Dim sql As String
    If bNovoEstado Then
        sql = "INSERT INTO IBS_Estado (IdEstado, UF, IBSUFpAliq, dIniVig, dFimVig) VALUES (" & sIdEstado & ",'" & sUF & "'," & sAliq & ",CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & ")"
    Else
        sql = "UPDATE IBS_Estado SET UF='" & sUF & "',IBSUFpAliq=" & sAliq & ",dIniVig=CONVERT(date,'" & sDIni & "',103),dFimVig=" & sDFimSQL & " WHERE Id=" & iEstadoId
    End If
    On Error GoTo ErrE
    vgDb.Execute sql
    MsgBox "Salvo.", vbInformation
    bNovoEstado = False: CarregarIBSEstado
    Exit Sub
ErrE: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluirEstado_Click()
    If iEstadoId = 0 Then Exit Sub
    If MsgBox("Excluir este registro de " & cboUFEdit.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrDE
    vgDb.Execute "DELETE FROM IBS_Estado WHERE Id=" & iEstadoId
    iEstadoId = 0: MsgBox "Exclu" & Chr(237) & "do.", vbInformation: CarregarIBSEstado
    Exit Sub
ErrDE: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdAplicarTodosUF_Click()
    Dim sAliq As String, sDIni As String, sDFim As String
    sAliq = Replace(Trim(txtAliqTodosUF.Text), ",", ".")
    sDIni = Trim(txtDIniTodosUF.Text)
    sDFim = Trim(txtDFimTodosUF.Text)
    If sAliq = "" Or sDIni = "" Then MsgBox "Preencha Al" & Chr(237) & "q e Vig. Ini.", vbExclamation: Exit Sub
    If MsgBox("Criar al" & Chr(237) & "quota " & txtAliqTodosUF.Text & "% para TODOS os estados?" & Chr(13) & "Vig: " & sDIni & " a " & IIf(sDFim = "", "sem fim", sDFim), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    On Error GoTo ErrAUF
    vgDb.Execute "INSERT INTO IBS_Estado (IdEstado, UF, IBSUFpAliq, dIniVig, dFimVig) SELECT DISTINCT IdEstado, UF, " & sAliq & ", CONVERT(date,'" & sDIni & "',103), " & sDFimSQL & " FROM Cidade WHERE IdEstado IS NOT NULL"
    MsgBox "Al" & Chr(237) & "quotas criadas para todos os estados.", vbInformation
    CarregarIBSEstado
    Exit Sub
ErrAUF: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ============================================================
' IBS MUNICIPIO
' ============================================================
Private Sub CarregarIBSMunicipio()
    lblCarregando.Visible = True
    lblCarregando.Caption = "CARREGANDO..."
    DoEvents
    Dim rRs As ADODB.Recordset
    Dim sql As String, r As Long, Filtro As String, nDots As Integer
    sql = "SELECT M.Id, C.UF, M.CodigoMunicipio, C.Nome, M.IBSMunpAliq, M.dIniVig, M.dFimVig " & _
          "FROM IBS_Municipio M " & _
          "JOIN (SELECT DISTINCT CAST(CodigoMunicipio AS NVARCHAR(7)) AS CodigoMunicipio, Nome, UF FROM Cidade) C " & _
          "  ON C.CodigoMunicipio = M.CodigoMunicipio"
    If cboFiltroUFMun.ListIndex > 0 Then Filtro = Filtro & " AND C.UF='" & cboFiltroUFMun.Text & "'"
    If Trim(txtFiltroCidade.Text) <> "" Then Filtro = Filtro & " AND C.Nome LIKE '%" & Trim(txtFiltroCidade.Text) & "%'"
    If Filtro <> "" Then sql = sql & " WHERE" & Mid(Filtro, 5)
    sql = sql & " ORDER BY C.UF, C.Nome, M.dIniVig"
    lstMunicipio.rows = 1
    RsOpen rRs, sql
    Do While Not rRs.EOF
        r = lstMunicipio.rows - 1
        lstMunicipio.AddItem ""
        lstMunicipio.Row = r: lstMunicipio.Col = 0: lstMunicipio.Text = rRs("UF")
        lstMunicipio.Row = r: lstMunicipio.Col = 1: lstMunicipio.Text = rRs("CodigoMunicipio")
        lstMunicipio.Row = r: lstMunicipio.Col = 2: lstMunicipio.Text = rRs("Nome")
        lstMunicipio.Row = r: lstMunicipio.Col = 3: lstMunicipio.Text = Format(rRs("IBSMunpAliq"), "0.00")
        lstMunicipio.Row = r: lstMunicipio.Col = 4: lstMunicipio.Text = IIf(IsNull(rRs("dIniVig")), "", Format(rRs("dIniVig"), "dd/mm/yyyy"))
        lstMunicipio.Row = r: lstMunicipio.Col = 5: lstMunicipio.Text = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "dd/mm/yyyy"))
        lstMunicipio.Row = r: lstMunicipio.Col = 6: lstMunicipio.Text = rRs("Id")
        If (lstMunicipio.rows - 1) Mod 30 = 0 Then
            nDots = (nDots Mod 3) + 1
            lblCarregando.Caption = "CARREGANDO" & String(nDots, ".")
            DoEvents
        End If
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    lblCarregando.Visible = False
End Sub

Private Sub lstMunicipio_Click()
    If lstMunicipio.Row = 0 Then Exit Sub
    iMunId = Val(lstMunicipio.TextMatrix(lstMunicipio.Row, 6))
    sCodMunSel = lstMunicipio.TextMatrix(lstMunicipio.Row, 1)
    txtCodMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 1)
    txtNomeMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 2)
    txtAliqMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 3)
    txtDIniMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 4)
    txtDFimMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 5)
    bNovoMun = False
End Sub

Private Sub cmdFiltrarMun_Click(): CarregarIBSMunicipio: End Sub
Private Sub cmdLimparMun_Click()
    cboFiltroUFMun.ListIndex = 0: txtFiltroCidade.Text = "": CarregarIBSMunicipio
End Sub

Private Sub cmdNovoMun_Click()
    If sCodMunSel = "" Then MsgBox "Selecione um munic" & Chr(237) & "pio na lista primeiro.", vbExclamation: Exit Sub
    bNovoMun = True: iMunId = 0
    txtAliqMunEdit.Text = "": txtDIniMunEdit.Text = "": txtDFimMunEdit.Text = ""
    txtAliqMunEdit.SetFocus
End Sub

Private Sub cmdSalvarMun_Click()
    Dim sAliq As String, sDIni As String, sDFim As String, sCod As String
    sCod = Trim(txtCodMunEdit.Text)
    sAliq = Replace(Trim(txtAliqMunEdit.Text), ",", ".")
    sDIni = Trim(txtDIniMunEdit.Text)
    sDFim = Trim(txtDFimMunEdit.Text)
    If sCod = "" Or sAliq = "" Or sDIni = "" Then
        MsgBox "Munic" & Chr(237) & "pio, Al" & Chr(237) & "q e Vig. Ini s" & Chr(227) & "o obrigat" & Chr(243) & "rios.", vbExclamation: Exit Sub
    End If
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    Dim sql As String
    If bNovoMun Then
        sql = "INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig) VALUES ('" & sCod & "'," & sAliq & ",CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & ")"
    Else
        sql = "UPDATE IBS_Municipio SET IBSMunpAliq=" & sAliq & ",dIniVig=CONVERT(date,'" & sDIni & "',103),dFimVig=" & sDFimSQL & " WHERE Id=" & iMunId
    End If
    On Error GoTo ErrM
    vgDb.Execute sql
    MsgBox "Salvo.", vbInformation
    bNovoMun = False: CarregarIBSMunicipio
    Exit Sub
ErrM: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluirMun_Click()
    If iMunId = 0 Then Exit Sub
    If MsgBox("Excluir este registro de " & txtNomeMunEdit.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrDM
    vgDb.Execute "DELETE FROM IBS_Municipio WHERE Id=" & iMunId
    iMunId = 0: MsgBox "Exclu" & Chr(237) & "do.", vbInformation: CarregarIBSMunicipio
    Exit Sub
ErrDM: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdAplicarTodosMun_Click()
    Dim sAliq As String, sDIni As String, sDFim As String
    sAliq = Replace(Trim(txtAliqTodosMun.Text), ",", ".")
    sDIni = Trim(txtDIniTodosMun.Text)
    sDFim = Trim(txtDFimTodosMun.Text)
    If sAliq = "" Or sDIni = "" Then MsgBox "Preencha Al" & Chr(237) & "q e Vig. Ini.", vbExclamation: Exit Sub
    Dim filtroWhere As String, descr As String
    If optEstadoFiltrado.Value And cboFiltroUFMun.ListIndex > 0 Then
        filtroWhere = " WHERE UF='" & cboFiltroUFMun.Text & "'"
        descr = "munic" & Chr(237) & "pios do estado " & cboFiltroUFMun.Text
    Else
        descr = "TODOS os munic" & Chr(237) & "pios"
    End If
    If MsgBox("Criar al" & Chr(237) & "quota " & txtAliqTodosMun.Text & "% para " & descr & "?" & Chr(13) & "Vig: " & sDIni & " a " & IIf(sDFim = "", "sem fim", sDFim), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    On Error GoTo ErrAM
    vgDb.Execute "INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig) SELECT DISTINCT CAST(CodigoMunicipio AS NVARCHAR(7))," & sAliq & ",CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & " FROM Cidade" & filtroWhere
    MsgBox "Al" & Chr(237) & "quotas criadas.", vbInformation: CarregarIBSMunicipio
    Exit Sub
ErrAM: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ============================================================
' SINCRONIZAR CIDADE
' ============================================================
Private Sub cmdAtualizarCidade_Click()
    Dim hoje As String
    hoje = Format(Now, "yyyy-mm-dd")
    On Error GoTo ErrSync
    vgDb.Execute "UPDATE C SET C.IBSUFpAliq=E.IBSUFpAliq FROM Cidade C JOIN IBS_Estado E ON C.IdEstado=E.IdEstado WHERE CONVERT(date,'" & hoje & "',23) BETWEEN E.dIniVig AND ISNULL(E.dFimVig,'9999-12-31')"
    vgDb.Execute "UPDATE C SET C.IBSMunpAliq=M.IBSMunpAliq FROM Cidade C JOIN IBS_Municipio M ON CAST(C.CodigoMunicipio AS NVARCHAR(7))=M.CodigoMunicipio WHERE CONVERT(date,'" & hoje & "',23) BETWEEN M.dIniVig AND ISNULL(M.dFimVig,'9999-12-31')"
    lblStatusSync.Caption = "Cidade atualizada em " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    lblAvisoSync.Caption = ""
    MsgBox "Tabela Cidade atualizada com al" & Chr(237) & "quotas vigentes de hoje.", vbInformation
    Exit Sub
ErrSync: MsgBox "Erro ao atualizar Cidade: " & Err.Description, vbCritical
End Sub

Private Sub VerificarAliquotasPendentes()
    Dim n As String
    n = SQLExecutaRetorno("SELECT CAST(COUNT(*) AS VARCHAR) AS Total FROM IBS_Estado WHERE dIniVig <= GETDATE() AND NOT EXISTS (SELECT 1 FROM Cidade C JOIN IBS_Estado E2 ON C.IdEstado=E2.IdEstado WHERE E2.Id=IBS_Estado.Id AND C.IBSUFpAliq=IBS_Estado.IBSUFpAliq)", "Total", "0")
    If Val(n) > 0 Then
        lblAvisoSync.Caption = "AVISO: Ha al" & Chr(237) & "quotas vigentes que ainda n" & Chr(227) & "o foram aplicadas em Cidade. Clique em 'Atualizar Cidade'."
    End If
End Sub
