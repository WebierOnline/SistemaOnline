VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Unidade_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE UNIDADES"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "Unidade_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "Unidade_Cadastro.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancelar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalvar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcluir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNovo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "Unidade_Cadastro.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         Height          =   4155
         Left            =   -74880
         ScaleHeight     =   4095
         ScaleWidth      =   7935
         TabIndex        =   15
         Top             =   420
         Width           =   7995
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3555
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   6271
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   4035
         Left            =   180
         ScaleHeight     =   3975
         ScaleWidth      =   6015
         TabIndex        =   13
         Top             =   480
         Width           =   6075
         Begin VB.TextBox txtResponsavel 
            Height          =   315
            Left            =   2340
            TabIndex        =   3
            Top             =   900
            Width           =   3615
         End
         Begin VB.ComboBox cboBairro 
            Height          =   315
            Left            =   60
            TabIndex        =   2
            Top             =   900
            Width           =   2235
         End
         Begin VB.TextBox txtUnidade 
            Height          =   315
            Left            =   60
            TabIndex        =   1
            Top             =   300
            Width           =   5895
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável:"
            Height          =   195
            Left            =   2340
            TabIndex        =   18
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   660
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unidade:"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   645
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   675
         Left            =   6360
         TabIndex        =   0
         Top             =   420
         Width           =   1755
         _ExtentX        =   3096
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
         MICON           =   "Unidade_Cadastro.frx":240A
         PICN            =   "Unidade_Cadastro.frx":2426
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
         Height          =   675
         Left            =   6360
         TabIndex        =   6
         Top             =   1140
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Alterar"
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
         MICON           =   "Unidade_Cadastro.frx":3100
         PICN            =   "Unidade_Cadastro.frx":311C
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
         Left            =   6360
         TabIndex        =   7
         Top             =   1860
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "Unidade_Cadastro.frx":39F6
         PICN            =   "Unidade_Cadastro.frx":3A12
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
         Height          =   675
         Left            =   6360
         TabIndex        =   4
         Top             =   1140
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
         MICON           =   "Unidade_Cadastro.frx":3D2C
         PICN            =   "Unidade_Cadastro.frx":3D48
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   675
         Left            =   6360
         TabIndex        =   5
         Top             =   1860
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
         MICON           =   "Unidade_Cadastro.frx":A612
         PICN            =   "Unidade_Cadastro.frx":A62E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   8265
      TabIndex        =   9
      Top             =   60
      Width           =   8295
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Unidade_Cadastro.frx":110D2
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDADES"
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
         Left            =   1095
         TabIndex        =   10
         Top             =   180
         Width           =   1590
      End
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   6960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   11
      Top             =   5595
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:26"
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
Attribute VB_Name = "Unidade_Cadastro"
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
   sSQL = "INSERT INTO unidade (codigo, unidade, bairro, responsavel) VALUES (" & _
      txtCodigo.Text & ", '" & txtUnidade.Text & "', '" & cboBairro.Text & "', '" & txtResponsavel.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE unidade SET unidade = '" & txtUnidade.Text & "', bairro = '" & cboBairro.Text & "', responsavel = '" & txtResponsavel.Text & "' WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_unidade FROM unidade;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then txtCodigo.Text = r("cod_unidade") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer, X As Integer
   
   With Grid
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 4500
     .ColWidth(3) = 2500
     .ColWidth(4) = 0
      
      For X = 0 To .Cols - 1
         .Col = X
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "UNIDADE"
      .TextMatrix(0, 3) = "BAIRRO"
      .TextMatrix(0, 4) = "BAIRRO"
      .Redraw = False
      
      i = 1
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("unidade")
            .TextMatrix(.Rows - 1, 3) = rTabela("bairro")
            .TextMatrix(.Rows - 1, 4) = rTabela("responsavel")
            rTabela.MoveNext
            
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Limpar_Objetos()
   txtCodigo.Text = ""
   txtUnidade.Text = ""
   cboBairro.Text = ""
   txtResponsavel.Text = ""
End Sub

Private Sub cboBairro_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varTexto As String
   varTexto = cboBairro.Text
   cboBairro.Clear
   
   
   sSQL = "SELECT DISTINCT bairro FROM unidade ORDER BY bairro;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboBairro.AddItem ValidateNull(r("bairro"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboBairro.Text = varTexto
   moCombo.AttachTo cboBairro
End Sub


Private Sub cboBairro_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Não informou os dados do registro
   If txtCodigo.Text = "" Or txtUnidade.Text = "" Then Exit Sub
   
   'Não é necessário consulta o registro antes de atualiza-lo
   'sSQL = "SELECT * FROM setor WHERE (cod_setor = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Form_Load
End Sub

Private Sub cmdCancelar_Click()
   Form_Load
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtCodigo.Text = "" Then Exit Sub
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir essa unidade?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM unidade WHERE (codigo = " & txtCodigo.Text & ");"
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
   cmdSalvar.Visible = True
   cmdCancelar.Visible = True
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   txtUnidade.SetFocus
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   
   'Não foi informado os dados do registro
   If txtCodigo.Text = "" Or txtUnidade.Text = "" Then Exit Sub
   
   'Não é necessário consultar todos os registros antes de inserir um novo
   'sSQL = "SELECT * FROM setor"
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
   Set moCombo = New cComboHelper
   MostrarGrid
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
   txtUnidade.Text = (Grid.TextMatrix(Grid.Row, 2))
   cboBairro.Text = (Grid.TextMatrix(Grid.Row, 3))
   txtResponsavel.Text = (Grid.TextMatrix(Grid.Row, 4))
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSalvar.Visible = False
   cmdCancelar.Visible = False
   SSTab1.Tab = 0
End Sub

Private Sub MostrarGrid()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT codigo, unidade, bairro, responsavel FROM unidade ORDER BY bairro;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Mostra os dados no grid
   FormatarGrid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtResponsavel_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtUnidade_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


