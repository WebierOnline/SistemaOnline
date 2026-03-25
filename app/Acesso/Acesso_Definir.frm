VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form Acesso_Definir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permissões de Usuário"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13110
   Icon            =   "Acesso_Definir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11865
      TabIndex        =   29
      Top             =   60
      Width           =   11895
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Acesso_Definir.frx":1172
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
         TabIndex        =   31
         Top             =   240
         Width           =   4080
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   7395
      Left            =   50
      ScaleHeight     =   7335
      ScaleWidth      =   11865
      TabIndex        =   4
      Top             =   1320
      Width           =   11925
      Begin VB.TextBox txtCodLogin 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3720
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   615
      End
      Begin VB.Frame Frame7 
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
         Height          =   2235
         Left            =   3060
         TabIndex        =   13
         Top             =   840
         Width           =   3435
         Begin VB.TextBox txtSenha3 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtSenha2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   1140
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtSenha1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   540
            Visible         =   0   'False
            Width           =   1455
         End
         Begin ChamaleonBtn.chameleonButton cmdSalvarSenha 
            Height          =   315
            Left            =   1860
            TabIndex        =   17
            Top             =   240
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
            MICON           =   "Acesso_Definir.frx":24CD
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
            Left            =   1860
            TabIndex        =   18
            Top             =   240
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
            MICON           =   "Acesso_Definir.frx":24E9
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
            Caption         =   "?"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblSenha2 
            Caption         =   "Confirmar"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   900
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblSenha1 
            Caption         =   "Senha (Até 8):"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   1740
            X2              =   1740
            Y1              =   240
            Y2              =   2100
         End
      End
      Begin VB.Frame frmTipoAcesso 
         Caption         =   "Tipo de Acesso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   60
         TabIndex        =   8
         Top             =   840
         Width           =   2955
         Begin VB.OptionButton optNivel4 
            Caption         =   "Somente PDV"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1515
         End
         Begin VB.OptionButton optNivel3 
            Caption         =   "Personalizado"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   960
            Width           =   1515
         End
         Begin VB.OptionButton optNivel2 
            Caption         =   "Limitado"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton optNivel1 
            Caption         =   "Controle Total"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Relatório"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9600
         TabIndex        =   7
         ToolTipText     =   "Emitir relatório de usuários cadastrados"
         Top             =   900
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3540
         TabIndex        =   2
         Top             =   3600
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Gravar Perfil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9540
         TabIndex        =   1
         ToolTipText     =   "Gravar perfil de acesso do usuário"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboLogin 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Usuários cadastrados"
         Top             =   360
         Width           =   3615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   75
         TabIndex        =   3
         Top             =   4200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   9660
         TabIndex        =   22
         Top             =   3240
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
         MICON           =   "Acesso_Definir.frx":2505
         PICN            =   "Acesso_Definir.frx":2521
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
         Left            =   9660
         TabIndex        =   23
         Top             =   3900
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
         MICON           =   "Acesso_Definir.frx":42B3
         PICN            =   "Acesso_Definir.frx":42CF
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
         Left            =   9660
         TabIndex        =   24
         Top             =   4560
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
         MICON           =   "Acesso_Definir.frx":6061
         PICN            =   "Acesso_Definir.frx":607D
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
         Left            =   9660
         TabIndex        =   25
         Top             =   2580
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
         MICON           =   "Acesso_Definir.frx":7E0F
         PICN            =   "Acesso_Definir.frx":7E2B
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
         Left            =   9660
         TabIndex        =   26
         Top             =   1920
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
         MICON           =   "Acesso_Definir.frx":9BBD
         PICN            =   "Acesso_Definir.frx":9BD9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimir 
         Height          =   555
         Left            =   10140
         TabIndex        =   27
         Top             =   5340
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Imprimir"
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
         MICON           =   "Acesso_Definir.frx":B96B
         PICN            =   "Acesso_Definir.frx":B987
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
         Left            =   10080
         TabIndex        =   28
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "Acesso_Definir.frx":D719
         PICN            =   "Acesso_Definir.frx":D735
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Left            =   180
         TabIndex        =   6
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o Usuário"
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
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Acesso_Definir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsUsuario As ADODB.Recordset
Public rsAcesso As ADODB.Recordset

Private Sub cboLogin_LostFocus()
   On Error GoTo TrataErro
   If cboLogin.Text = "" Then txtCodLogin.Text = "": Exit Sub
   
   txtCodLogin = cboLogin.ItemData(cboLogin.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub Check1_Click()
'--------------------------
'Marcar/Desmarcar os perfis
'--------------------------
If Check1.Value = 1 Then
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(i).Checked = True
    Next
Else
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(i).Checked = False
    Next
End If
End Sub

Private Sub cmdSalvarSenha_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtSenha1.Text = "" Or txtSenha2.Text = "" Or txtCodigo.Text = "" Then Exit Sub
   
   If txtSenha1.Text <> txtSenha2.Text Then
      ShowMsg "As senhas digitas são diferentes!", vbInformation
      txtSenha1.Text = ""
      txtSenha2.Text = ""
      txtSenha1.SetFocus
      Exit Sub
   End If
   
   dbData.Execute "UPDATE usuario SET senha = '" & txtSenha1.Text & "' WHERE (codigo = " & txtCodigo.Text & ");"

   'VERIFICAR SE A SENHA FOI CADASTRADA E EXIBE OS DADOS
   sSQL = "SELECT senha FROM usuario WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If IsNull(r("senha")) Then
      txtSenha1.Visible = True
      txtSenha2.Visible = True
      txtSenha3.Visible = False
      lblSenha1.Visible = True
      lblSenha1.Caption = "Senha"
      lblSenha2.Visible = True
      lblSenha2.Caption = "Confirmar"
      lblSenha3.Visible = False
      lblSenha3.Caption = "?"
      cmdSalvarSenha.Visible = True
      cmdAlterarSenha.Visible = False
   Else
      txtSenha1.Visible = True
      txtSenha1.Text = ""
      txtSenha2.Visible = True
      txtSenha2.Text = ""
      txtSenha3.Visible = True
      txtSenha3.Text = ""
      lblSenha1.Visible = True
      lblSenha1.Caption = "Senha Atual"
      lblSenha2.Visible = True
      lblSenha2.Caption = "Nova Senha"
      lblSenha3.Visible = True
      lblSenha3.Caption = "Confirmar"
      cmdSalvarSenha.Visible = False
      cmdAlterarSenha.Visible = True
   End If
End Sub


Private Sub cmdAlterarSenha_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim var_Senha As String
   
   If txtSenha2.Text = "" Or txtSenha3.Text = "" Or txtCodigo.Text = "" Then Exit Sub
   
   'VERIFICAR SE A SENHA ATUAL ESTÁ ERRADA
   sSQL = "SELECT senha FROM usuario WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then var_Senha = r("senha")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   If var_Senha <> txtSenha1.Text Then
      ShowMsg "A Senha Atual não está correta!", vbInformation
      txtSenha1.Text = ""
      txtSenha1.SetFocus
      Exit Sub
   End If
   
   'VERIFICAR SE AS SENHAS FORAM DIGITADAS IGUAIS
   If txtSenha2.Text <> txtSenha3.Text Then
      ShowMsg "As senhas digitas são diferentes!", vbInformation
      txtSenha2.Text = ""
      txtSenha3.Text = ""
      txtSenha2.SetFocus
      Exit Sub
   End If
   
   dbData.Execute "UPDATE usuario SET senha = '" & txtSenha2.Text & "' WHERE (codigo = " & txtCodigo.Text & ");"
   txtSenha1.Text = ""
   txtSenha2.Text = ""
   txtSenha3.Text = ""
End Sub


Private Sub cbologin_Click()
'---------------------------------------
'Mostrar o perfil do usuário selecionado
'---------------------------------------
rsUsuario.MoveFirst
rsUsuario.Find "login='" & cboLogin & "'"
Check1.Value = 0

If Not rsUsuario.EOF Then
    Call LerAcesso
Else
    MsgBox "Ocorreu um erro ao localizar usuário!", vbExclamation, "Acesso"
End If

cboLogin_LostFocus
End Sub

Private Sub Command1_Click()
If MsgBox("Confirma a alteração no perfil de acesso?", vbQuestion + vbYesNo + vbDefaultButton2, "Controle de Acesso") = vbYes Then
    
    rsUsuario.MoveFirst
    rsUsuario.Find "login='" & cboLogin & "'"
    
    If Not rsUsuario.EOF Then
       Call GravarAcesso
    Else
        MsgBox "Ocorreu um erro ao localizar usuário!", vbExclamation, "Acesso"
    End If
End If
End Sub


Private Sub Command2_Click()
'----------------
'Emitir relatório
'----------------
If MsgBox("Confirma a emissão do relatório?", vbQuestion + vbYesNo + vbDefaultButton2, "Relatório") = vbYes Then
    Dim Imagem As RptImage
    Set Imagem = DataReport1.Sections("Section4").Controls("image1")
    Set Imagem.Picture = LoadPicture(App.path & "\vbmania_logo.jpg")
    Set DataReport1.DataSource = rsUsuario
    DataReport1.Show
End If
End Sub

Private Sub Form_Activate()
ListarUsuario
ListarAcesso
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call Desconectar
End Sub

Private Sub optNivel1_Click()
   If txtCodigo.Text = "" Then Exit Sub
   dbData.Execute "UPDATE usuario SET nivel = " & 1 & " WHERE (codigo = " & txtCodigo.Text & ");"
End Sub

'Private Sub optNivel1_Click()
'   If txtCodigo.Text = "" Then Exit Sub
'   dbData.Execute "UPDATE usuario SET nivel = " & 1 & " WHERE (codigo = " & txtCodigo.Text & ");"
'End Sub

Private Sub optNivel2_Click()
   If txtCodigo.Text = "" Then Exit Sub
   dbData.Execute "UPDATE usuario SET nivel = " & 2 & " WHERE (codigo = " & txtCodigo.Text & ");"
End Sub

Private Sub optNivel3_Click()
   If txtCodigo.Text = "" Then Exit Sub
   dbData.Execute "UPDATE usuario SET nivel = " & 3 & " WHERE (codigo = " & txtCodigo.Text & ");"
End Sub

Private Sub optNivel4_Click()
   If txtCodigo.Text = "" Then Exit Sub
   dbData.Execute "UPDATE usuario SET nivel = " & 4 & " WHERE (codigo = " & txtCodigo.Text & ");"
End Sub

Public Sub ListarAcesso()
Dim sSQL As String

sSQL = "SELECT * FROM Usuario_permissoes ORDER BY codigo"
Set rsAcesso = dbData.OpenRecordset(sSQL)

Do While Not rsAcesso.EOF
    ListView1.ListItems.Add , , rsAcesso.Fields(2)
    rsAcesso.MoveNext
Loop
End Sub

Public Sub ListarUsuario()
Dim sSQL As String

sSQL = "SELECT *, login, codigo FROM usuario ORDER BY codigo"
Set rsUsuario = dbData.OpenRecordset(sSQL)

Do While Not rsUsuario.EOF
    cboLogin.AddItem rsUsuario.Fields("login")
    cboLogin.ItemData(cboLogin.NewIndex) = rsUsuario("codigo")
    rsUsuario.MoveNext
Loop
End Sub

Sub GravarAcesso()
For i = 1 To ListView1.ListItems.Count

'rsUsuario("cliinc") = 1
        If ListView1.ListItems.Item(i).Text = "Clientes - Inclusão" Then
            rsUsuario.Fields("cliinc") = IIf(ListView1.ListItems.Item(i).Checked = True, "1", "0")
        ElseIf ListView1.ListItems.Item(i).Text = "Clientes - Alteração" Then
            rsUsuario.Fields("clialt") = IIf(ListView1.ListItems.Item(i).Checked = True, "1", "0")
        ElseIf ListView1.ListItems.Item(i).Text = "Clientes - Exclusão" Then
            rsUsuario.Fields("cliexc") = IIf(ListView1.ListItems.Item(i).Checked = True, "1", "0")
        ElseIf ListView1.ListItems.Item(i).Text = "Produtos - Inclusão" Then
            rsUsuario.Fields("prodinc") = IIf(ListView1.ListItems.Item(i).Checked = True, "1", "0")
        ElseIf ListView1.ListItems.Item(i).Text = "Produtos - Alteração" Then
            rsUsuario.Fields("prodalt") = IIf(ListView1.ListItems.Item(i).Checked = True, "1", "0")
        ElseIf ListView1.ListItems.Item(i).Text = "Produtos - Exclusão" Then
            rsUsuario.Fields("prodexc") = IIf(ListView1.ListItems.Item(i).Checked = True, "1", "0")
        End If
Next
rsUsuario.Update
MsgBox "Perfil de acesso cadastrado!", vbInformation, "Perfil"
End Sub

Sub LerAcesso()
'------------------------------------------
'Ler os acessos configurados para o usuário
'------------------------------------------

For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = "Clientes - Inclusão" Then
            ListView1.ListItems.Item(i).Checked = IIf(rsUsuario.Fields("cliinc") = 1, True, False)
        ElseIf ListView1.ListItems.Item(i).Text = "Clientes - Alteração" Then
            ListView1.ListItems.Item(i).Checked = IIf(rsUsuario.Fields("clialt") = 1, True, False)
        ElseIf ListView1.ListItems.Item(i).Text = "Clientes - Exclusão" Then
            ListView1.ListItems.Item(i).Checked = IIf(rsUsuario.Fields("cliexc") = 1, True, False)
        ElseIf ListView1.ListItems.Item(i).Text = "Produtos - Inclusão" Then
            ListView1.ListItems.Item(i).Checked = IIf(rsUsuario.Fields("prodinc") = 1, True, False)
        ElseIf ListView1.ListItems.Item(i).Text = "Produtos - Alteração" Then
            ListView1.ListItems.Item(i).Checked = IIf(rsUsuario.Fields("prodalt") = 1, True, False)
        ElseIf ListView1.ListItems.Item(i).Text = "Produtos - Exclusão" Then
            ListView1.ListItems.Item(i).Checked = IIf(rsUsuario.Fields("prodexc") = 1, True, False)
        End If
Next
End Sub
