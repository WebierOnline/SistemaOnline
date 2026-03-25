VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Usuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permissões de Usuário"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   Icon            =   "Usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   10605
      TabIndex        =   14
      Top             =   60
      Width           =   10635
      Begin VB.TextBox txtCodUsuario 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   9240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   8940
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Usuarios.frx":1172
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PERMISSÕES DE USUÁRIO"
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
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   4080
      End
   End
   Begin VB.PictureBox frmLogin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   60
      ScaleHeight     =   5265
      ScaleWidth      =   8295
      TabIndex        =   2
      Top             =   3060
      Width           =   8325
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   3780
         TabIndex        =   30
         Top             =   2100
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cboVisivel 
         Height          =   315
         Left            =   6120
         TabIndex        =   28
         Top             =   420
         Width           =   2115
      End
      Begin VB.TextBox cboLogin 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         TabIndex        =   27
         Top             =   420
         Width           =   3075
      End
      Begin VB.ComboBox cboNivel 
         Height          =   315
         Left            =   3240
         TabIndex        =   25
         Top             =   420
         Width           =   2835
      End
      Begin VB.TextBox txtCodNivel 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5460
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCodLogin 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame frmSenha 
         Caption         =   "Senha"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   8115
         Begin VB.TextBox txtSenha3 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3120
            MaxLength       =   8
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   540
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtSenha2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   540
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtSenha1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   540
            Visible         =   0   'False
            Width           =   1455
         End
         Begin ChamaleonBtn.chameleonButton cmdSalvarSenha 
            Height          =   315
            Left            =   3120
            TabIndex        =   9
            Top             =   540
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Salvar"
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
            MICON           =   "Usuarios.frx":24CD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAlterarSenha 
            Height          =   315
            Left            =   4620
            TabIndex        =   10
            Top             =   540
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MICON           =   "Usuarios.frx":24E9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblSenha3 
            Caption         =   "Repetir"
            Height          =   195
            Left            =   3120
            TabIndex        =   13
            Top             =   300
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblSenha2 
            Caption         =   "Confirmar"
            Height          =   195
            Left            =   1620
            TabIndex        =   12
            Top             =   300
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblSenha1 
            Caption         =   "Senha (Até 6):"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Marcar/Desmarcar todas as operações"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   0
         Top             =   1860
         Width           =   3735
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   2100
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   14235
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visivel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   29
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nível"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operações realizadas no Sistema"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1860
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   750
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   615
      Left            =   8460
      TabIndex        =   20
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
      MICON           =   "Usuarios.frx":2505
      PICN            =   "Usuarios.frx":2521
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
      Left            =   8520
      TabIndex        =   21
      Top             =   7740
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "Usuarios.frx":42B3
      PICN            =   "Usuarios.frx":42CF
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
      Left            =   8460
      TabIndex        =   22
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
      MICON           =   "Usuarios.frx":6061
      PICN            =   "Usuarios.frx":607D
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
      TabIndex        =   23
      Top             =   8385
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14631
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:52"
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1935
      Left            =   60
      TabIndex        =   26
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin ChamaleonBtn.chameleonButton chameleonButton1 
      Height          =   315
      Left            =   9060
      TabIndex        =   31
      Top             =   5340
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
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
      MICON           =   "Usuarios.frx":7E0F
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
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private printSQL As String
Dim i As Integer
Dim idUsuario As Long
Dim itemAtual As String
Dim itemSaida As String
Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub Limpar_Objetos()
cboLogin.Text = ""
txtCodLogin.Text = ""
cboNivel.Text = ""
txtCodNivel.Text = ""
txtSenha1.Text = ""
txtSenha2.Text = ""
txtSenha3.Text = ""
cboVisivel.Text = ""
txtCodUsuario.Text = ""
End Sub

Private Sub Mostrar_Permissoes()
'Marca/Desmarca permissões já definidas do usuario
If txtCodUsuario.Text = "" Then Exit Sub

idUsuario = txtCodUsuario.Text

For i = 1 To ListView1.ListItems.Count
    If SQLExecutaRetorno("SELECT COUNT(*) r FROM usuario_acessos WHERE Cod_Usuario = " & idUsuario & " AND Cod_Permissao = " & Mid(ListView1.ListItems.Item(i).key, 2), "r", 0) > 0 Then
       ListView1.ListItems.Item(i).Checked = True
    Else
       ListView1.ListItems.Item(i).Checked = False
    End If
Next
End Sub

Private Sub cboLogin_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboNivel_GotFocus()
Dim codAtual As String

itemAtual = cboNivel.Text
codAtual = txtCodNivel.Text

cboNivel.Clear
cboNivel.AddItem "1 - Total"
cboNivel.AddItem "2 - Limitado"
cboNivel.AddItem "3 - Somente PDV"
cboNivel.AddItem "4 - Personalizado"

cboNivel.Text = itemAtual
txtCodNivel.Text = codAtual

moCombo.AttachTo cboNivel
End Sub


Private Sub cboNivel_LostFocus()
Dim idUsuario As Long

If cboNivel.Text = "" Then txtCodNivel.Text = "": Exit Sub

itemSaida = cboNivel.Text

idUsuario = Grid.TextMatrix(Grid.Row, 1)

If itemSaida <> itemAtual Then
    sSQL = "DELETE FROM Usuario_Acessos WHERE Cod_Usuario = " & idUsuario
    vgDb.Execute sSQL
End If

If cboNivel.Text = "1 - Total" Then
    txtCodNivel.Text = "1"
    sSQL = "DELETE FROM Usuario_Acessos WHERE Cod_Usuario = " & idUsuario
    vgDb.Execute sSQL
    sSQL = "INSERT INTO Usuario_Acessos (Cod_Usuario, Cod_Permissao) " & _
           "SELECT " & idUsuario & ", Codigo FROM Usuario_permissoes"
    vgDb.Execute sSQL
ElseIf cboNivel.Text = "2 - Limitado" Then
    txtCodNivel.Text = "2"
ElseIf cboNivel.Text = "3 - Somente PDV" Then
    txtCodNivel.Text = "3"
ElseIf cboNivel.Text = "4 - Personalizado" Then
    txtCodNivel.Text = "4"
End If

Call Mostrar_Permissoes
End Sub




Private Sub cboVisivel_GotFocus()
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboVisivel.Text

cboVisivel.Clear
cboVisivel.AddItem "SIM"
cboVisivel.AddItem "NÃO"

cboVisivel.Text = itemAtual

moCombo.AttachTo cboVisivel
End Sub


Private Sub chameleonButton1_Click()
'exibir dados no ListView
sSQL = "SELECT Usuario_Acessos.Cod_Usuario, Usuario_Acessos.Cod_Permissao, Usuario_permissoes.permissao " & _
"FROM usuario INNER JOIN usuario_acessos ON usuario.codigo = Usuario_Acessos.cod_usuario INNER JOIN usuario_permissoes ON usuario_permissoes.codigo = Usuario_Acessos.cod_permissao " & _
"WHERE (usuario_acessos.cod_usuario = " & txtCodLogin.Text & ")"

Set r = dbData.OpenRecordset(sSQL)

 If r.BOF Then
    If r.State <> 0 Then r.Close
    Set r = Nothing
    Exit Sub
 End If

Check1.Value = 0

Dim i As Integer

If ListView1.ListItems(i).ListSubItems(1).Text = "ESTORNAR" Then
    ListView1.ListItems.Item(i).Checked = True
ElseIf ListView1.ListItems(i).ListSubItems(1).Text = "AJUSTE" Then
    ListView1.ListItems.Item(i).Checked = True
End If
End Sub

Private Sub cmdAlterarSenha_Click()
cmdSalvarSenha_Click
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos
frmLogin.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub


Private Sub cmdSalvar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodUsuario.Text = "" Or cboLogin.Text = "" Then Exit Sub

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdSalvarSenha_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtSenha1.Text = "" Or txtSenha2.Text = "" Or txtCodUsuario.Text = "" Then Exit Sub

If txtCodUsuario.Text = "" Then Exit Sub

sSQL = "SELECT codigo, password FROM usuario WHERE (codigo = " & txtCodUsuario.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Exit Sub
End If
    
If Not r Is Nothing Then
    If IsNull(r("password")) Then
        If txtSenha1.Text <> txtSenha2.Text Then
            ShowMsg "As senhas digitas são diferentes!", vbInformation
            txtSenha1.Text = ""
            txtSenha2.Text = ""
            txtSenha1.SetFocus
            Exit Sub
        End If
        dbData.Execute "UPDATE usuario SET password = '" & txtSenha1.Text & "' WHERE (codigo = " & txtCodUsuario.Text & ");"
    Else
    
        If r("password") <> txtSenha1.Text Then
        MsgBox r("password")
            ShowMsg "Senha Atual errada!", vbInformation
            txtSenha1.Text = ""
            Exit Sub
        End If
        
        If txtSenha2.Text <> txtSenha3.Text Then
            ShowMsg "As senhas digitas são diferentes!", vbInformation
            txtSenha2.Text = ""
            txtSenha3.Text = ""
            txtSenha2.SetFocus
            Exit Sub
        End If
            
        dbData.Execute "UPDATE usuario SET password = '" & txtSenha3.Text & "' WHERE (codigo = " & txtCodUsuario.Text & ");"
    End If

txtSenha1.Visible = True
txtSenha2.Visible = True
txtSenha3.Visible = False
lblSenha1.Visible = True
lblSenha1.Caption = "Senha(até 8)"
lblSenha2.Visible = True
lblSenha2.Caption = "Confirmar"
lblSenha3.Visible = False
lblSenha3.Caption = "?"
cmdSalvarSenha.Visible = True
cmdAlterarSenha.Visible = False
frmSenha.Enabled = False
txtSenha1.Text = ""
txtSenha2.Text = ""
txtSenha3.Text = ""
End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
MostrarGrid
'MostrarPermissoes
ListarAcesso
'StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmLogin.Enabled = False
End Sub

Public Sub ListarAcesso()
Dim i As Integer
Dim sSQL As String
Dim rs As ADODB.Recordset


sSQL = "SELECT codigo, permissao from usuario_permissoes order by permissao;"
Set rs = dbData.OpenRecordset(sSQL)

ListView1.ListItems.Clear

Do While Not rs.EOF
    ListView1.ListItems.Add , "z" + Trim(str$(rs.Fields(0))), rs.Fields(1)
    rs.MoveNext
Loop
End Sub
Private Sub MostrarPermissoes()
Dim NewItem As ListItem
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * from usuario_permissoes;"
Set rs = dbData.OpenRecordset(sSQL)

If Not rs.EOF Then
    Call LerAcesso
End If


'If r.State <> 0 Then r.Close
'Set r = Nothing
End Sub
Sub LerAcesso()
Dim i As Integer

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Text = "ESTORNAR" Then
        ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 1, True, False)
    ElseIf ListView1.ListItems.Item(i).Text = "AJUSTE" Then
        ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 2, True, False)
    ElseIf ListView1.ListItems.Item(i).Text = "EXCLUIR" Then
        ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 3, True, False)
    ElseIf ListView1.ListItems.Item(i).Text = "REATIVAR" Then
        ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 4, True, False)
    End If
Next

'For i = 1 To ListView1.ListItems.Count
'    If ListView1.ListItems.Item(i).Text = "ESTORNAR" Then ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 1, True, False)

'    If ListView1.ListItems.Item(i).Text = "AJUSTE" Then ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 2, True, False)

'    If ListView1.ListItems.Item(i).Text = "EXCLUIR" Then ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 3, True, False)

'    If ListView1.ListItems.Item(i).Text = "REATIVAR" Then ListView1.ListItems.Item(i).Checked = IIf(r.Fields("Cod_Permissao") = 4, True, False)
'Next

'If ListView1.ListItems(i).ListSubItems(3).Text = "ESTORNAR" Then
'    ListView1.ListItems.Item(i).Checked = True
'ElseIf ListView1.ListItems(i).ListSubItems(0).Text = "AJUSTE" Then
'    ListView1.ListItems.Item(i).Checked = True
'End If

End Sub

Private Sub MostrarGrid()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT codigo, login, nivel,  CASE WHEN password IS NULL THEN 'NÃO' ELSE 'SIM' END as var_senha, CASE visivel WHEN 0 THEN 'NÃO' ELSE 'SIM' END as var_visivel from usuario order by codigo;"
Set r = dbData.OpenRecordset(sSQL)

'Mostra os dados no grid
FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer

With Grid
   .Clear
   .Cols = 6
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 800
   .ColWidth(2) = 2500
   .ColWidth(3) = 1000
   .ColWidth(4) = 1000
   .ColWidth(5) = 1000
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   
   .TextMatrix(0, 1) = "COD"
   .TextMatrix(0, 2) = "LOGIN"
   .TextMatrix(0, 3) = "SENHA"
   .TextMatrix(0, 4) = "NIVEL"
   .TextMatrix(0, 5) = "EXIBIR"
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("codigo")
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("login"))
         .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_senha"))
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("nivel"))
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_visivel"))
         rTabela.MoveNext
         
         .rows = .rows + 1
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
   
   .rows = .rows - 1
   .Redraw = True
End With
End Sub
Private Function Atualizar_Dados() As Boolean
Dim sSQL As String

Dim varVisivel As Boolean
   
If cboVisivel.Text = "SIM" Then
    varVisivel = True
ElseIf cboVisivel.Text = "NÃO" Then
    varVisivel = False
End If

sSQL = "UPDATE USUARIO SET nivel = '" & txtCodNivel.Text & "', visivel = '" & Abs(varVisivel) & "'  WHERE (codigo = " & txtCodUsuario.Text & ");"

Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Grid_Click()
txtCodLogin.Text = ""
txtCodUsuario.Text = (Grid.TextMatrix(Grid.Row, 1))

Call Mostrar_Permissoes

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim idUsuario As Long
   
   idUsuario = Grid.TextMatrix(Grid.Row, 1)
   
   If idUsuario = 0 Then Exit Sub
   
   If Item.Checked Then
      sSQL = "INSERT INTO [Usuario_Acessos] ([Cod_Usuario],[Cod_Permissao]) VALUES " & _
             "(" & idUsuario & ", " & Mid(Item.key, 2) & ")"
      vgDb.Execute sSQL
   Else
      sSQL = "DELETE FROM Usuario_Acessos WHERE Cod_Usuario = " & idUsuario & " AND Cod_Permissao = " & Mid(Item.key, 2)
      vgDb.Execute sSQL
   End If
End Sub

Private Sub txtCodLogin_Change()
If txtCodLogin.Text = "" Then Exit Sub

'If cmdExcluir.Enabled = True Then
   sSQL = "SELECT codigo, login, password FROM Usuario WHERE (codigo = " & txtCodLogin.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
    If r.BOF Then
       If r.State <> 0 Then r.Close
       Set r = Nothing
       Exit Sub
    End If

If txtCodLogin.Text <> "" Then
    frmSenha.Enabled = True
Else
    MsgBox "Ocorreu um erro ao localizar usuário!", vbExclamation, "Acesso"
    frmSenha.Enabled = False
End If
    
Mostrar_Dados r

'exibir dados no ListView
sSQL = "SELECT Usuario_Acessos.Cod_Usuario, Usuario_Acessos.Cod_Permissao, Usuario_permissoes.permissao " & _
"FROM usuario INNER JOIN usuario_acessos ON usuario.codigo = Usuario_Acessos.cod_usuario INNER JOIN usuario_permissoes ON usuario_permissoes.codigo = Usuario_Acessos.cod_permissao " & _
"WHERE (usuario_acessos.cod_usuario = " & txtCodLogin.Text & ")"

Set r = dbData.OpenRecordset(sSQL)

 If r.BOF Then
    If r.State <> 0 Then r.Close
    Set r = Nothing
    Exit Sub
 End If

Check1.Value = 0

'Call LerAcesso
End Sub


Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)

If Not rTabela Is Nothing Then
'txtAgencia.Text = ValidateNull(rTabela("agencia"))
'txtConta.Text = ValidateNull(rTabela("conta"))
'cboTipoConta.Text = ValidateNull(rTabela("tipo_conta"))

'If Abs(CInt(rTabela("acesso"))) = 1 Then
'   chkSistema.Value = Checked
'ElseIf Abs(CInt(rTabela("acesso"))) = 0 Then
'   chkSistema.Value = Unchecked
'End If

'If chkSistema.Value = Checked Then Frame7.Enabled = True

'If IsNull(rTabela("estornar")) = False Then
'    If Abs(CInt(rTabela("estornar"))) = 1 Then
'       chkEstornar.Value = Checked
'    ElseIf Abs(CInt(rTabela("estornar"))) = 0 Then
'       chkEstornar.Value = Unchecked
'    End If
'Else
'    chkEstornar.Value = Unchecked
'End If

'If Abs(CInt(rTabela("ativo"))) = 1 Then
'   chkInativo.Value = Checked
'ElseIf Abs(CInt(rTabela("ativo"))) = 0 Then
'   chkInativo.Value = Unchecked
'End If

'      chkSistema.Value = IIf(CInt(rTabela("acesso")) = 1, 1, 0)
'      chkInativo.Value = Abs(CInt(rTabela("ativo")))

'If Not IsNull(rTabela("nivel")) Then
'   If rTabela("nivel") = 1 Then
      'optNivel1.Value = True
'   ElseIf rTabela("nivel") = 2 Then
      'optNivel2.Value = True
'   ElseIf rTabela("nivel") = 3 Then
      'optNivel3.Value = True
'   ElseIf rTabela("nivel") = 4 Then
      'optNivel4.Value = True
'   End If
'End If

If IsNull(rTabela("password")) Then
   txtSenha1.Visible = True
   txtSenha2.Visible = True
   txtSenha3.Visible = False
   lblSenha1.Visible = True
   lblSenha1.Caption = "Senha (Até 6):"
   lblSenha2.Visible = True
   lblSenha2.Caption = "Confirmar"
   lblSenha3.Visible = False
   lblSenha3.Caption = "?"
   cmdSalvarSenha.Visible = True
   cmdAlterarSenha.Visible = False
Else
   txtSenha1.Visible = True
   txtSenha2.Visible = True
   txtSenha3.Visible = True
   lblSenha1.Visible = True
   lblSenha1.Caption = "Senha Atual"
   lblSenha2.Visible = True
   lblSenha2.Caption = "Nova Senha(até 6)"
   lblSenha3.Visible = True
   lblSenha3.Caption = "Confirmar"
   cmdSalvarSenha.Visible = False
   cmdAlterarSenha.Visible = True
End If

End If
End Sub

Private Sub txtCodNivel_Change()
If txtCodNivel.Text = "" Then cboNivel.Text = "": Exit Sub

If txtCodNivel.Text = 1 Then
    cboNivel.Text = "1 - Total"
ElseIf txtCodNivel.Text = 2 Then
    cboNivel.Text = "2 - Limitado"
ElseIf txtCodNivel.Text = 3 Then
    cboNivel.Text = "3 - Somente PDV"
ElseIf txtCodNivel.Text = 4 Then
    cboNivel.Text = "4 - Personalizado"
End If
End Sub

Private Sub txtCodUsuario_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodUsuario.Text = "" Then Exit Sub

'If cmdAlterar.Enabled = True Then
   sSQL = "SELECT *, (CASE WHEN visivel = 1 THEN 'SIM' ELSE 'NÃO' END) as VarVisivel FROM usuario WHERE (codigo = " & txtCodUsuario.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then Mostrar_Usuario r
   If r.State <> 0 Then r.Close
   Set r = Nothing
'End If

'cboSetor.Text = (Grid.TextMatrix(Grid.Row, 2))
'cmdAlterar.Visible = True
'cmdExcluir.Visible = True
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
frmLogin.Enabled = True
End Sub


Private Sub Mostrar_Usuario(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
   cboLogin.Text = rTabela("login")
   txtCodLogin.Text = rTabela("codigo")
   txtCodNivel.Text = ValidateNull(rTabela("nivel"))
   cboVisivel.Text = rTabela("VarVisivel")
End If
End Sub

