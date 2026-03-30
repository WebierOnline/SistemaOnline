VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Carros_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "Carros_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      ScaleHeight     =   1065
      ScaleWidth      =   7125
      TabIndex        =   12
      Top             =   60
      Width           =   7155
      Begin VB.Image imgCarro 
         Height          =   900
         Left            =   300
         Picture         =   "Carros_Cadastro.frx":23D2
         Top             =   60
         Width           =   1065
      End
      Begin VB.Image imgMoto 
         Height          =   1185
         Left            =   120
         Picture         =   "Carros_Cadastro.frx":2C1A
         Top             =   -60
         Width           =   1395
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CADASTRO DE MOTOS"
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
         Left            =   1830
         TabIndex        =   13
         Top             =   300
         Width           =   3555
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   5235
      Left            =   3300
      ScaleHeight     =   5175
      ScaleWidth      =   3855
      TabIndex        =   10
      Top             =   1200
      Width           =   3915
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   60
         TabIndex        =   15
         Top             =   0
         Width           =   3735
         Begin VB.OptionButton optModelo 
            Caption         =   "Modelo"
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
            Left            =   1620
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optFabricante 
            Caption         =   "Fabricante"
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
            Left            =   180
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4515
         Left            =   60
         TabIndex        =   14
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7964
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSalvar 
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   6540
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Salvar"
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
      MICON           =   "Carros_Cadastro.frx":4146
      PICN            =   "Carros_Cadastro.frx":4162
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox frmPrincipal 
      Enabled         =   0   'False
      Height          =   5235
      Left            =   60
      ScaleHeight     =   5175
      ScaleWidth      =   3135
      TabIndex        =   3
      Top             =   1200
      Width           =   3195
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2340
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtModelo 
         Height          =   315
         Left            =   900
         MaxLength       =   40
         TabIndex        =   2
         Top             =   540
         Width           =   2175
      End
      Begin VB.ComboBox cboFabricante 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   675
      Left            =   1500
      TabIndex        =   7
      Top             =   6540
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
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
      MICON           =   "Carros_Cadastro.frx":52AC
      PICN            =   "Carros_Cadastro.frx":52C8
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
      Height          =   675
      Left            =   2940
      TabIndex        =   8
      Top             =   6540
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
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
      MICON           =   "Carros_Cadastro.frx":5BA2
      PICN            =   "Carros_Cadastro.frx":5BBE
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
      Height          =   675
      Left            =   4380
      TabIndex        =   0
      Top             =   6540
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "Carros_Cadastro.frx":5ED8
      PICN            =   "Carros_Cadastro.frx":5EF4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   675
      Left            =   5820
      TabIndex        =   9
      Top             =   6540
      Width           =   1395
      _ExtentX        =   2461
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
      MICON           =   "Carros_Cadastro.frx":6BCE
      PICN            =   "Carros_Cadastro.frx":6BEA
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
      TabIndex        =   18
      Top             =   7335
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8573
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:18"
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
Attribute VB_Name = "Carros_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Private Function Inserir_Dados() As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO carros (cod_carro, fabricante, modelo) VALUES (" & txtCodigo.Text & ", '" & cboFabricante.Text & "', '" & txtModelo.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE carros SET " & _
      "fabricante = '" & cboFabricante.Text & "', " & _
      "modelo = '" & txtModelo.Text & "' "
   
   'Condição para atualização
   sSQL = sSQL & "WHERE (cod_carro = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(cod_carro), 0) AS codigo FROM carros;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("codigo") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Limpar_Objetos()
   txtCodigo.Text = ""
   cboFabricante.Text = ""
   txtModelo.Text = ""
End Sub

Private Sub PreencherGrid()
   Dim INDICE As String
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If optFabricante.Value = True Then
      INDICE = "fabricante"
   ElseIf optModelo.Value = True Then
      INDICE = "modelo"
   End If
   
   sSQL = "SELECT * FROM carros ORDER BY " & INDICE
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cboFabricante_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboFabricante.Clear
   
   sSQL = "SELECT fabricante FROM carros GROUP BY fabricante;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboFabricante.AddItem r("fabricante")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboFabricante
End Sub

Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Se os campos não estão preenchidos
   If txtCodigo.Text = "" Or cboFabricante.Text = "" Or txtModelo.Text = "" Then Exit Sub
   
   'Não é necessário consulta o registro antes de atualiza-lo
   'sSQL = "SELECT * FROM carros WHERE (cod_carro = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Form_Load
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtCodigo.Text = "" Then Exit Sub
   
   'Solicita confirmação do usuário
   If ShowMsg("Excluir essa MODELO?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Não é necessário consulta o registro antes de exclui-lo
   'sSQL = "SELECT * FROM carros WHERE (cod_carro = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM carros WHERE (cod_carro = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos
   Form_Load
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdNovo_Click()
   frmPrincipal.Enabled = True
   Limpar_Objetos
   Auto_Numeracao
   cmdSalvar.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cboFabricante.SetFocus
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   
   'Se os dados do form não forma preenchidos, sai da rotina
   If txtCodigo.Text = "" Or cboFabricante.Text = "" Or txtModelo.Text = "" Then Exit Sub
   
   'Não é necessário consultar todos os registros antes de inserir um novo
   'sSQL = "SELECT * FROM carros;"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'A auto numeração do código deve ser utilizada no momento de salvar o registro
   'para evitar duplicidade de código para quando houver mais de um terminal operando
   'ao mesmo tempo
   'AutoNumeracao
   
   'Faz a inserção de forma direta e verifica se houve algum erro
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
   Dim oCfg As ConfigItem
   
   PreencherGrid
   Set oCfg = sysConfig("TIPO_OS")
   
   'sSQL = "SELECT * FROM configuracao WHERE (codigo = 1);"
   'Set r = dbData.OpenRecordset(sSQL)
   
   If UCase(oCfg.Value) = "CARROS" Then
      imgMoto.Visible = False
      imgCarro.Visible = True
      lblTitulo.Caption = "CADASTRO DE AUTOMOVEIS"
   ElseIf UCase(oCfg.Value) = "MOTOS" Then
      imgMoto.Visible = True
      imgCarro.Visible = False
      lblTitulo.Caption = "CADASTRO DE MOTOS"
   ElseIf UCase(oCfg.Value) = "INFOR" Then
      imgMoto.Visible = False
      imgCarro.Visible = False
      lblTitulo.Caption = ""
   Else
      imgMoto.Visible = False
      imgCarro.Visible = False
      lblTitulo.Caption = ""
   End If
   
   cmdSalvar.Enabled = False
   cmdExcluir.Enabled = False
   cmdAlterar.Enabled = False
   frmPrincipal.Enabled = False
   
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   frmPrincipal.Enabled = True
   cmdExcluir.Enabled = True
   cmdAlterar.Enabled = True
   cmdSalvar.Enabled = False
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
   cboFabricante.Text = (Grid.TextMatrix(Grid.Row, 2))
   txtModelo.Text = (Grid.TextMatrix(Grid.Row, 3))
End Sub

Private Sub optFabricante_Click()
   PreencherGrid
End Sub

Private Sub optModelo_Click()
   PreencherGrid
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 4
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1500
      .ColWidth(3) = 1900
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "FABRICANTE"
      .TextMatrix(0, 3) = "MODELO"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("cod_carro")
            .TextMatrix(.Rows - 1, 2) = rTabela("fabricante")
            .TextMatrix(.Rows - 1, 3) = rTabela("modelo")
            
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
