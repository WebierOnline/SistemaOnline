VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form OS_Recapadora_Pneus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE PNEUS"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "OS_Recapadora_Pneus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   7965
      TabIndex        =   9
      Top             =   60
      Width           =   7995
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7020
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CADASTRO DE PNEUS"
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
         Left            =   1470
         TabIndex        =   10
         Top             =   300
         Width           =   3450
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   120
         Picture         =   "OS_Recapadora_Pneus.frx":23D2
         Top             =   45
         Width           =   900
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4755
      Left            =   60
      TabIndex        =   7
      Top             =   1140
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   8387
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
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "OS_Recapadora_Pneus.frx":28A9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSair"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNovo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalvar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExcluir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAlterar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "OS_Recapadora_Pneus.frx":28C5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   11
         Top             =   420
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   6376
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.PictureBox Picture1 
         Enabled         =   0   'False
         Height          =   4155
         Left            =   120
         ScaleHeight     =   4095
         ScaleWidth      =   5355
         TabIndex        =   8
         Top             =   420
         Width           =   5415
         Begin VB.TextBox txtMedida 
            Height          =   315
            Left            =   2940
            TabIndex        =   15
            Top             =   300
            Width           =   2355
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   60
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   300
            Width           =   2835
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medida"
            Height          =   195
            Left            =   2940
            TabIndex        =   16
            Top             =   60
            Width           =   525
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   315
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   5640
         TabIndex        =   3
         Top             =   1740
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "OS_Recapadora_Pneus.frx":28E1
         PICN            =   "OS_Recapadora_Pneus.frx":28FD
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
         Left            =   5640
         TabIndex        =   4
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "OS_Recapadora_Pneus.frx":468F
         PICN            =   "OS_Recapadora_Pneus.frx":46AB
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
         Left            =   5640
         TabIndex        =   5
         Top             =   3060
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "OS_Recapadora_Pneus.frx":643D
         PICN            =   "OS_Recapadora_Pneus.frx":6459
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
         Left            =   5640
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "OS_Recapadora_Pneus.frx":81EB
         PICN            =   "OS_Recapadora_Pneus.frx":8207
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
         Left            =   5640
         TabIndex        =   0
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "OS_Recapadora_Pneus.frx":9F99
         PICN            =   "OS_Recapadora_Pneus.frx":9FB5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   5640
         TabIndex        =   6
         Top             =   3960
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1085
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
         MICON           =   "OS_Recapadora_Pneus.frx":BD47
         PICN            =   "OS_Recapadora_Pneus.frx":BD63
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   12
      Top             =   5955
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10028
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "13:09"
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
Attribute VB_Name = "OS_Recapadora_Pneus"
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
   sSQL = "INSERT INTO OS_Recapadora_Pneus (codigo, tipo, medida) VALUES (" & txtCodigo.Text & ", '" & cboTipo.Text & "', '" & txtMedida.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

'Comando de atualização
sSQL = "UPDATE OS_Recapadora_Pneus SET tipo = '" & cboTipo.Text & "', medida = '" & txtMedida.Text & "' WHERE (codigo = " & txtCodigo.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_servico FROM OS_Recapadora_Pneus;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then txtCodigo.Text = r("cod_servico") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid
   .Clear
   .Cols = 4
   .Rows = 2
   
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 2000
    .ColWidth(3) = 4000
   
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next i
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "TIPO"
   .TextMatrix(0, 3) = "MEDIDA"
   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.Rows - 1, 2) = rTabela("tipo")
         .TextMatrix(.Rows - 1, 3) = rTabela("MEDIDA")
         
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
cboTipo.Text = ""
txtMedida.Text = ""
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
If rTabela Is Nothing Then Exit Sub
txtCodigo.Text = ValidateNull(rTabela("codigo"))
cboTipo.Text = ValidateNull(rTabela("tipo"))
txtMedida.Text = ValidateNull(rTabela("medida"))
End Sub

Private Sub cboTipo_GotFocus()
Dim itemAtual As String
itemAtual = cboTipo.Text
cboTipo.Clear
cboTipo.AddItem "AGRÍCOLA"
cboTipo.AddItem "CARGA"
cboTipo.AddItem "CAMINHONETE"
cboTipo.Text = itemAtual
moCombo.AttachTo cboTipo
End Sub


Private Sub cboTipo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cmdAlterar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

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

If ShowMsg("Excluir esse serviço?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

sSQL = "DELETE FROM OS_Recapadora_Pneus WHERE (codigo = " & txtCodigo.Text & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdNovo_Click()
Picture1.Enabled = True
Limpar_Objetos
Auto_Numeracao
cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cboTipo.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo TrataErro

If txtCodigo.Text = "" Or cboTipo.Text = "" Or txtMedida.Text = "" Then Exit Sub

If Not Inserir_Dados Then
    ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
    Exit Sub
End If

Limpar_Objetos
Form_Load
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
Set moCombo = New cComboHelper

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
Picture1.Enabled = False
MostrarGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
Picture1.Enabled = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
cboTipo.Text = (Grid.TextMatrix(Grid.Row, 2))
txtMedida.Text = (Grid.TextMatrix(Grid.Row, 3))

cmdNovo.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
SSTab1.Tab = 0
cboTipo.SetFocus
End Sub

Private Sub MostrarGrid()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * FROM OS_Recapadora_Pneus ORDER BY MEDIDA;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
