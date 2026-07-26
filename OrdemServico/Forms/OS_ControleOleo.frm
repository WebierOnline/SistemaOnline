VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form OS_ControleOleo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIMITE DE PRODUTOS"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmAlterar 
      Caption         =   "Limites"
      Height          =   1035
      Left            =   7740
      TabIndex        =   11
      Top             =   5280
      Width           =   3855
      Begin VB.OptionButton optKM 
         Caption         =   "KM"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
      Begin VB.OptionButton optPrazo 
         Caption         =   "Prazo (meses)"
         Height          =   255
         Left            =   900
         TabIndex        =   13
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txtValor 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2220
      End
      Begin ChamaleonBtn.chameleonButton cmdAtualizar 
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Alterar"
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
         MICON           =   "OS_ControleOleo.frx":0000
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
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1035
      Left            =   60
      TabIndex        =   2
      Top             =   5280
      Width           =   7635
      Begin VB.Frame frmCriterio 
         Caption         =   "Critérios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         TabIndex        =   7
         Top             =   120
         Width           =   6255
         Begin VB.ComboBox cboCriterio 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   420
            Width           =   5175
         End
         Begin ChamaleonBtn.chameleonButton cmdFiltrar 
            Height          =   315
            Left            =   5340
            TabIndex        =   10
            Top             =   420
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Buscar"
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
            MICON           =   "OS_ControleOleo.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblxxx 
            AutoSize        =   -1  'True
            Caption         =   "lblxxx"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   200
            Width           =   375
         End
      End
      Begin VB.OptionButton optCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   975
      End
      Begin VB.OptionButton optDescricao 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optCodBarra 
         Caption         =   "Cód. Barra"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   1095
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   795
      End
   End
   Begin MSFlexGridLib.MSFlexGrid lstOleo 
      Height          =   4860
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   8573
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Image ImgMarcada 
      Height          =   195
      Left            =   4620
      Picture         =   "OS_ControleOleo.frx":0038
      Top             =   4980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcada 
      Height          =   195
      Left            =   4920
      Picture         =   "OS_ControleOleo.frx":2437
      Top             =   4980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcadaTODAS 
      Height          =   195
      Left            =   60
      Picture         =   "OS_ControleOleo.frx":47B3
      Top             =   4980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImgMarcadaTODAS 
      Height          =   195
      Left            =   60
      Picture         =   "OS_ControleOleo.frx":6B2F
      Top             =   4980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMarcarTodas 
      AutoSize        =   -1  'True
      Caption         =   "Marcar Todas"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   4980
      Width           =   990
   End
End
Attribute VB_Name = "OS_ControleOleo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bAutoCompletando As Boolean

Private Sub Form_Load()
    ConfigurarGrid
    optTodos.Value = True

    optKM.Value = True
    optPrazo.Value = False
    txtValor.Enabled = True
End Sub

Private Sub ConfigurarGrid()
    With lstOleo
        .Cols = 6
        .Rows = 1
        .FixedCols = 0
        .ColWidth(0) = 360
        .ColWidth(1) = 0
        .ColWidth(2) = 1500
        .ColWidth(3) = 5600
        .ColWidth(4) = 1400
        .ColWidth(5) = 1400
        .TextMatrix(0, 2) = "CÓD. BARRA"
        .TextMatrix(0, 3) = "ÓLEO E CIA"
        .TextMatrix(0, 4) = "KM TROCA"
        .TextMatrix(0, 5) = "TEMPO TROCA"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
End Sub

Private Sub CarregarGrid()
    Dim r As ADODB.Recordset
    Dim sql As String
    Dim sWhere As String
    Dim n As Integer
    Dim lChk As Long

    sWhere = MontarFiltro()
    If sWhere = "" Then Exit Sub

    lstOleo.Rows = 1

    sql = "SELECT produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
          "OS_ControleOleo.LIMITE_KM AS var_km, OS_ControleOleo.LIMITE_PRAZO AS var_prazo " & _
          "FROM produtos LEFT JOIN OS_ControleOleo ON OS_ControleOleo.COD_PRODUTO = produtos.codigo " & _
          "WHERE " & sWhere & " " & _
          "ORDER BY produtos.descricao"

    RsOpen r, sql
    n = 1
    Do While Not r.EOF
        With lstOleo
            .Rows = n + 1
            .Row = n: .Col = 0: .Text = ""
            .Row = n: .Col = 1: .Text = ValidateNull(r("var_cod"))
            .Row = n: .Col = 2: .Text = ValidateNull(r("var_codbarra"))
            .Row = n: .Col = 3: .Text = ValidateNull(r("var_desc"))
            .Row = n: .Col = 4: .Text = ValidateNull(r("var_km"))
            .Row = n: .Col = 5: .Text = ValidateNull(r("var_prazo"))
        End With
        n = n + 1
        r.MoveNext
    Loop
    If r.State <> 0 Then r.Close

    For lChk = 1 To lstOleo.Rows - 1
        lstOleo.Row = lChk
        lstOleo.Col = 0
        Set lstOleo.CellPicture = imgDesmarcada.Picture
        lstOleo.CellPictureAlignment = 4
    Next lChk

    imgDesmarcadaTODAS.Visible = True
    ImgMarcadaTODAS.Visible = False
    lblMarcarTodas.Caption = "Marcar Todas"
End Sub

Private Function MontarFiltro() As String
    Dim sBase As String

    sBase = "(produtos.descricao LIKE '%OLEO%' OR produtos.descricao LIKE '%ÓLEO%') AND (produtos.ativo = 1)"

    If optCodBarra.Value = True Then
        If Trim(cboCriterio.Text) = "" Then
            MsgBox "Digite ou selecione o código de barra!", vbInformation, "Aviso do Sistema"
            Exit Function
        End If
        MontarFiltro = sBase & " AND (produtos.cod_barra = '" & Replace(Trim(cboCriterio.Text), "'", "''") & "')"
    ElseIf optDescricao.Value = True Then
        If Trim(cboCriterio.Text) = "" Then
            MsgBox "Digite ou selecione a descrição!", vbInformation, "Aviso do Sistema"
            Exit Function
        End If
        MontarFiltro = sBase & " AND (produtos.descricao LIKE '%" & Replace(Trim(cboCriterio.Text), "'", "''") & "%')"
    ElseIf optCategoria.Value = True Then
        If Trim(cboCriterio.Text) = "" Then
            MsgBox "Selecione a categoria!", vbInformation, "Aviso do Sistema"
            Exit Function
        End If
        MontarFiltro = sBase & " AND (produtos.categoria = '" & Replace(Trim(cboCriterio.Text), "'", "''") & "')"
    Else
        MontarFiltro = sBase
    End If
End Function

Private Sub cmdFiltrar_Click()
    CarregarGrid
End Sub

Private Sub optTodos_Click()
    cboCriterio.Enabled = False
    cboCriterio.Visible = False
    cmdFiltrar.Visible = False
    cboCriterio.Clear
    cboCriterio.Text = ""
    lblxxx.Caption = ""
    CarregarGrid
End Sub

Private Sub optCodBarra_Click()
    lblxxx.Caption = "Cód. Barra:"
    cboCriterio.Enabled = True
    cboCriterio.Visible = True
    cmdFiltrar.Visible = True
    PreencherCriterioCodBarra
    If Me.Visible Then cboCriterio.SetFocus
End Sub

Private Sub optDescricao_Click()
    lblxxx.Caption = "Descrição:"
    cboCriterio.Enabled = True
    cboCriterio.Visible = True
    cmdFiltrar.Visible = True
    PreencherCriterioDescricao
    If Me.Visible Then cboCriterio.SetFocus
End Sub

Private Sub optCategoria_Click()
    lblxxx.Caption = "Categoria:"
    cboCriterio.Enabled = True
    cboCriterio.Visible = True
    cmdFiltrar.Visible = True
    PreencherCriterioCategoria
    If Me.Visible Then cboCriterio.SetFocus
End Sub

Private Sub PreencherCriterioCodBarra()
    Dim r As ADODB.Recordset
    Dim sql As String

    cboCriterio.Clear
    cboCriterio.Text = ""

    sql = "SELECT DISTINCT produtos.cod_barra FROM produtos " & _
          "WHERE (produtos.descricao LIKE '%OLEO%' OR produtos.descricao LIKE '%ÓLEO%') AND (produtos.ativo = 1) " & _
          "AND (produtos.cod_barra IS NOT NULL) AND (produtos.cod_barra <> '') " & _
          "ORDER BY produtos.cod_barra"

    RsOpen r, sql
    Do While Not r.EOF
        cboCriterio.AddItem ValidateNull(r("cod_barra"))
        r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
End Sub

Private Sub PreencherCriterioDescricao()
    Dim r As ADODB.Recordset
    Dim sql As String

    cboCriterio.Clear
    cboCriterio.Text = ""

    sql = "SELECT produtos.codigo, produtos.descricao FROM produtos " & _
          "WHERE (produtos.descricao LIKE '%OLEO%' OR produtos.descricao LIKE '%ÓLEO%') AND (produtos.ativo = 1) " & _
          "ORDER BY produtos.descricao"

    RsOpen r, sql
    Do While Not r.EOF
        cboCriterio.AddItem ValidateNull(r("descricao"))
        cboCriterio.ItemData(cboCriterio.NewIndex) = r("codigo")
        r.MoveNext
    Loop
    If r.State <> 0 Then r.Close
End Sub

Private Sub PreencherCriterioCategoria()
    cboCriterio.Clear
    cboCriterio.Text = ""
    cboCriterio.AddItem "ÓLEOS"
    cboCriterio.AddItem "FILTROS"
    cboCriterio.AddItem "LUBRIFICANTES"
    cboCriterio.AddItem "FLUIDOS"
    cboCriterio.AddItem "ADITIVOS"
End Sub

Private Sub cboCriterio_Change()
    Dim sTexto As String
    Dim sItem As String
    Dim i As Integer

    If bAutoCompletando Then Exit Sub
    If Not optDescricao.Value Then Exit Sub

    sTexto = cboCriterio.Text

    If Len(sTexto) < 3 Then Exit Sub

    For i = 0 To cboCriterio.ListCount - 1
        sItem = cboCriterio.List(i)
        If UCase$(Left$(sItem, Len(sTexto))) = UCase$(sTexto) Then
            bAutoCompletando = True
            cboCriterio.Text = sItem
            cboCriterio.SelStart = Len(sTexto)
            cboCriterio.SelLength = Len(sItem) - Len(sTexto)
            bAutoCompletando = False
            Exit For
        End If
    Next i
End Sub

Private Sub lstOleo_Click()
    Dim iRow As Long

    If lstOleo.Col <> 0 Then Exit Sub
    iRow = lstOleo.Row
    If iRow < 1 Then Exit Sub

    If lstOleo.TextMatrix(iRow, 0) = "1" Then
        lstOleo.TextMatrix(iRow, 0) = ""
        lstOleo.Row = iRow: lstOleo.Col = 0
        Set lstOleo.CellPicture = imgDesmarcada.Picture
    Else
        lstOleo.TextMatrix(iRow, 0) = "1"
        lstOleo.Row = iRow: lstOleo.Col = 0
        Set lstOleo.CellPicture = ImgMarcada.Picture
    End If
    lstOleo.CellPictureAlignment = 4
End Sub

Private Sub imgDesmarcadaTODAS_Click()
    Dim i As Long

    imgDesmarcadaTODAS.Visible = False
    ImgMarcadaTODAS.Visible = True
    lblMarcarTodas.Caption = "Desmarcar Todas"

    For i = 1 To lstOleo.Rows - 1
        lstOleo.TextMatrix(i, 0) = "1"
        lstOleo.Row = i: lstOleo.Col = 0
        Set lstOleo.CellPicture = ImgMarcada.Picture
        lstOleo.CellPictureAlignment = 4
    Next i
End Sub

Private Sub ImgMarcadaTODAS_Click()
    Dim i As Long

    ImgMarcadaTODAS.Visible = False
    imgDesmarcadaTODAS.Visible = True
    lblMarcarTodas.Caption = "Marcar Todas"

    For i = 1 To lstOleo.Rows - 1
        lstOleo.TextMatrix(i, 0) = ""
        lstOleo.Row = i: lstOleo.Col = 0
        Set lstOleo.CellPicture = imgDesmarcada.Picture
        lstOleo.CellPictureAlignment = 4
    Next i
End Sub

Private Sub optKM_Click()
    txtValor.Enabled = True
    txtValor.Text = ""
    If Me.Visible Then txtValor.SetFocus
End Sub

Private Sub optPrazo_Click()
    txtValor.Enabled = True
    txtValor.Text = ""
    If Me.Visible Then txtValor.SetFocus
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = aNumeros(KeyAscii)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAtualizar_Click
    End If
End Sub

Private Sub cmdAtualizar_Click()
    Dim i As Long
    Dim lCodProduto As Long
    Dim lValor As Long
    Dim rChk As ADODB.Recordset
    Dim bAlgumMarcado As Boolean

    If Trim(txtValor.Text) = "" Or Not IsNumeric(txtValor.Text) Then
        MsgBox "Digite um valor numérico válido!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    lValor = Val(txtValor.Text)

    If optPrazo.Value = True And lValor > 36 Then
        MsgBox "O prazo não pode ser maior que 36 meses!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    bAlgumMarcado = False

    For i = 1 To lstOleo.Rows - 1
        If lstOleo.TextMatrix(i, 0) = "1" Then
            bAlgumMarcado = True
            lCodProduto = Val(lstOleo.TextMatrix(i, 1))

            Set rChk = dbData.OpenRecordset("SELECT ID FROM OS_ControleOleo WHERE (COD_PRODUTO = " & lCodProduto & ")")
            If rChk.EOF Then
                If optKM.Value = True Then
                    dbData.Execute "INSERT INTO OS_ControleOleo (COD_PRODUTO, LIMITE_KM) VALUES (" & lCodProduto & ", " & lValor & ")"
                Else
                    dbData.Execute "INSERT INTO OS_ControleOleo (COD_PRODUTO, LIMITE_PRAZO) VALUES (" & lCodProduto & ", " & lValor & ")"
                End If
            Else
                If optKM.Value = True Then
                    dbData.Execute "UPDATE OS_ControleOleo SET LIMITE_KM = " & lValor & " WHERE (COD_PRODUTO = " & lCodProduto & ")"
                Else
                    dbData.Execute "UPDATE OS_ControleOleo SET LIMITE_PRAZO = " & lValor & " WHERE (COD_PRODUTO = " & lCodProduto & ")"
                End If
            End If
            rChk.Close

            If optKM.Value = True Then
                lstOleo.TextMatrix(i, 4) = lValor
            Else
                lstOleo.TextMatrix(i, 5) = lValor
            End If
        End If
    Next i

    If Not bAlgumMarcado Then
        MsgBox "Nenhum produto marcado no grid!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    MsgBox "Atualizado com sucesso!", vbInformation, "Aviso do Sistema"
    txtValor.Text = ""
    ImgMarcadaTODAS_Click
End Sub

