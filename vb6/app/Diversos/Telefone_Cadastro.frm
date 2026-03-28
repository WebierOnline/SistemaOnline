VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Telefone_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TELEFONES"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "Telefone_Cadastro.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      ScaleHeight     =   1065
      ScaleWidth      =   7245
      TabIndex        =   24
      Top             =   60
      Width           =   7275
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   0
         Picture         =   "Telefone_Cadastro.frx":23D2
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CADASTRO"
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
         Left            =   1740
         TabIndex        =   25
         Top             =   360
         Width           =   1785
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox frmCadastro 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   60
      ScaleHeight     =   2805
      ScaleWidth      =   7245
      TabIndex        =   16
      Top             =   1200
      Width           =   7275
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1020
         Width           =   3555
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   7035
      End
      Begin VB.Frame frmCelular 
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         TabIndex        =   19
         Top             =   1440
         Width           =   3075
         Begin VB.ComboBox cboCelularOP2 
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   660
            Width           =   1155
         End
         Begin VB.ComboBox cboCelularOP1 
            Height          =   315
            Left            =   60
            TabIndex        =   6
            Top             =   300
            Width           =   1155
         End
         Begin MSMask.MaskEdBox mskCelular1 
            Height          =   315
            Left            =   1260
            TabIndex        =   7
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCelular2 
            Height          =   315
            Left            =   1260
            TabIndex        =   9
            Top             =   660
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame frmResidencial 
         Caption         =   "Residencial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2100
         TabIndex        =   18
         Top             =   1440
         Width           =   1875
         Begin MSMask.MaskEdBox mskResidencial1 
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskResidencial2 
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   660
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame frmComercial 
         Caption         =   "Comercial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1875
         Begin VB.CheckBox optCOM0800 
            Caption         =   "0800"
            Height          =   195
            Left            =   1080
            TabIndex        =   22
            Top             =   0
            Width           =   675
         End
         Begin MSMask.MaskEdBox mskComercial1 
            Height          =   315
            Left            =   60
            TabIndex        =   2
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskComercial2 
            Height          =   315
            Left            =   60
            TabIndex        =   3
            Top             =   660
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   555
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   615
      Left            =   7380
      TabIndex        =   12
      Top             =   3420
      Width           =   1755
      _ExtentX        =   3096
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
      MICON           =   "Telefone_Cadastro.frx":2C31
      PICN            =   "Telefone_Cadastro.frx":2C4D
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
      Left            =   7380
      TabIndex        =   10
      Top             =   780
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Telefone_Cadastro.frx":2F67
      PICN            =   "Telefone_Cadastro.frx":2F83
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   615
      Left            =   7380
      TabIndex        =   13
      Top             =   2100
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
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
      MICON           =   "Telefone_Cadastro.frx":984D
      PICN            =   "Telefone_Cadastro.frx":9869
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
      Height          =   615
      Left            =   7380
      TabIndex        =   11
      Top             =   1440
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Telefone_Cadastro.frx":A143
      PICN            =   "Telefone_Cadastro.frx":A15F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluir 
      Height          =   615
      Left            =   7380
      TabIndex        =   14
      Top             =   2760
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
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
      MICON           =   "Telefone_Cadastro.frx":10C03
      PICN            =   "Telefone_Cadastro.frx":10C1F
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
      Left            =   7380
      TabIndex        =   23
      Top             =   60
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "Telefone_Cadastro.frx":10F39
      PICN            =   "Telefone_Cadastro.frx":10F55
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
      TabIndex        =   26
      Top             =   4125
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:30"
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
Attribute VB_Name = "Telefone_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Private Function Inserir_Dados() As Boolean
   Dim sSQL As String
   
   sSQL = "INSERT INTO telefone (codigo, nome, email, com0800, comercial1, residencial1, celular1, comercial2, " & _
      "residencial2, celular2, op1, op2) VALUES (" & txtCodigo.Text & ", '" & txtNome.Text & "', '" & txtEmail.Text & "', " & _
      IIf(optCOM0800.Value = 1, 1, 0) & ", '" & mskComercial1.Text & "', '" & mskResidencial1.Text & "', '" & _
      mskCelular1.Text & "', '" & mskComercial2.Text & "', '" & mskResidencial2.Text & "', '" & mskCelular2.Text & "', '" & _
      cboCelularOP1.Text & "', '" & cboCelularOP2.Text & "');"
   
   Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   sSQL = "UPDATE telefone SET " & _
      "nome = '" & txtNome.Text & "', " & _
      "email = '" & txtEmail.Text & "', " & _
      "com0800 = " & IIf(optCOM0800.Value = 1, 1, 0) & ", " & _
      "comercial1 = '" & mskComercial1.Text & "', " & _
      "residencial1 = '" & mskResidencial1.Text & "', " & _
      "celular1 = '" & mskCelular1.Text & "', " & _
      "comercial2 = '" & mskComercial2.Text & "', " & _
      "residencial2 = '" & mskResidencial2.Text & "', " & _
      "celular2 = '" & mskCelular2.Text & "', " & _
      "op1 = '" & cboCelularOP1.Text & "', " & _
      "op2 = '" & cboCelularOP2.Text & "' "
   
   sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"
   
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub AutoNumeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_telefone FROM telefone;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("cod_telefone") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

'Rotina para preenchimento
Sub PreencherOperadora(objLista As ComboBox)
   objLista.Clear
   objLista.AddItem "VIVO"
   objLista.AddItem "CLARO"
   objLista.AddItem "TIM"
   objLista.AddItem "OI"
End Sub

Private Sub cboCelularOP1_GotFocus()
   PreencherOperadora cboCelularOP1
   moCombo.AttachTo cboCelularOP1
End Sub

Private Sub cboCelularOP2_GotFocus()
   PreencherOperadora cboCelularOP2
   moCombo.AttachTo cboCelularOP2
End Sub

Private Sub cmdCancelar_Click()
   Campos_Brancos
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   'cmdSair.Enabled = False
   frmCadastro.Enabled = False
End Sub

Private Sub cmdNovo_Click()
   Campos_Brancos
   AutoNumeracao
   cmdNovo.Enabled = False
   cmdSalvar.Enabled = True
   cmdCancelar.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   'cmdSair.Enabled = False
   frmCadastro.Enabled = True
   txtNome.SetFocus
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtNome.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   If mskComercial1.Text = "" And mskResidencial1.Text = "" And mskCelular1.Text = "" And mskComercial2.Text = "" And mskResidencial2.Text = "" And mskCelular2.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   'sSQL = "SELECT * FROM telefone WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   sSQL = "DELETE FROM telefone WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbExclamation
      Exit Sub
   End If
   
   Campos_Brancos
   Form_Load
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtNome.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   If mskComercial1.Text = "" And mskResidencial1.Text = "" And mskCelular1.Text = "" And mskComercial2.Text = "" And mskResidencial2.Text = "" And mskCelular2.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   'sSQL = "SELECT * FROM telefone WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Campos_Brancos
   Form_Load
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   cmdNovo.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   'cmdSair.Enabled = False
   frmCadastro.Enabled = False
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Set moCombo = New cComboHelper
End Sub

Private Sub Campos_Brancos()
   If cmdAlterar.Enabled <> True Then txtCodigo.Text = ""
   txtNome.Text = ""
   txtEmail.Text = ""
   mskComercial1.Mask = ""
   mskComercial1.Text = ""
   mskComercial2.Mask = ""
   mskComercial2.Text = ""
   mskResidencial1.Mask = ""
   mskResidencial1.Text = ""
   mskResidencial2.Mask = ""
   mskResidencial2.Text = ""
   mskCelular1.Mask = ""
   mskCelular1.Text = ""
   mskCelular2.Mask = ""
   mskCelular2.Text = ""
   cboCelularOP1.Text = ""
   cboCelularOP2.Text = ""
   optCOM0800.Value = False
   frmCadastro.Enabled = False
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      txtCodigo.Text = rTabela("codigo")
      txtNome.Text = rTabela("nome")
      txtEmail.Text = rTabela("email")
      mskComercial1.Text = rTabela("comercial1")
      mskResidencial1.Text = rTabela("residencial1")
      mskCelular1.Text = rTabela("celular1")
      
      If rTabela("com0800") = 1 Then
         optCOM0800.Value = 1
      Else
         optCOM0800.Value = 0
      End If
   End If
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   
   If txtNome.Text = "" Then Exit Sub
   
   If mskComercial1.Text = "" And mskResidencial1.Text = "" And mskCelular1.Text = "" And mskComercial2.Text = "" And mskResidencial2.Text = "" And mskCelular2.Text = "" Then Exit Sub
   
   'Faz a inserção de forma direta e verifica se houve algum erro
   If Not Inserir_Dados Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Campos_Brancos
   Form_Load
   Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se este recado já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub mskCelular1_GotFocus()
   mskCelular1.SelStart = 1
   If cmdAlterar.Enabled = True Then SelectControl mskCelular1
End Sub

Private Sub mskCelular1_KeyPress(KeyAscii As Integer)
   mskCelular1.Mask = "(##) ####-####"
End Sub

Private Sub mskCelular1_LostFocus()
   If mskCelular1.Text = "(__) ____-____" Then
      mskCelular1.Mask = ""
      mskCelular1.Text = ""
   End If
End Sub

Private Sub mskCelular2_GotFocus()
   mskCelular2.SelStart = 1
   If cmdAlterar.Enabled = True Then SelectControl mskCelular2
End Sub

Private Sub mskCelular2_KeyPress(KeyAscii As Integer)
   mskCelular2.Mask = "(##) ####-####"
End Sub

Private Sub mskCelular2_LostFocus()
   If mskCelular2.Text = "(__) ____-____" Then
      mskCelular2.Mask = ""
      mskCelular2.Text = ""
   End If
End Sub

Private Sub mskComercial1_GotFocus()
   If optCOM0800.Value = 1 Then
      mskComercial1.SelStart = 3
   Else
      mskComercial1.SelStart = 2
   End If
   
   If cmdAlterar.Enabled = True Then SelectControl mskComercial1
End Sub

Private Sub mskComercial1_KeyPress(KeyAscii As Integer)
   If optCOM0800.Value = 1 Then
      mskComercial1.Mask = "(0#00) ###-####"
   Else
      mskComercial1.Mask = "(##) ####-####"
   End If
End Sub

Private Sub mskComercial1_LostFocus()
   If mskComercial1.Text = "(__) ____-____" Then
      mskComercial1.Mask = ""
      mskComercial1.Text = ""
   ElseIf mskComercial1.Text = "(0_00) ___-____" Then
      mskComercial1.Mask = ""
      mskComercial1.Text = ""
   End If
End Sub

Private Sub mskComercial2_GotFocus()
   mskComercial2.SelStart = 1
   If cmdAlterar.Enabled = True Then SelectControl mskComercial2
End Sub

Private Sub mskComercial2_KeyPress(KeyAscii As Integer)
   mskComercial2.Mask = "(##) ####-####"
End Sub

Private Sub mskComercial2_LostFocus()
   If mskComercial2.Text = "(__) ____-____" Then
      mskComercial2.Mask = ""
      mskComercial2.Text = ""
   End If
End Sub

Private Sub mskResidencial1_GotFocus()
   mskResidencial1.SelStart = 1
   If cmdAlterar.Enabled = True Then SelectControl mskResidencial1
End Sub

Private Sub mskResidencial1_KeyPress(KeyAscii As Integer)
   mskResidencial1.Mask = "(##) ####-####"
End Sub

Private Sub mskResidencial1_LostFocus()
   If mskResidencial1.Text = "(__) ____-____" Then
      mskResidencial1.Mask = ""
      mskResidencial1.Text = ""
   End If
End Sub

Private Sub mskResidencial2_GotFocus()
   mskResidencial2.SelStart = 1
   If cmdAlterar.Enabled = True Then SelectControl mskResidencial2
End Sub

Private Sub mskResidencial2_KeyPress(KeyAscii As Integer)
    mskResidencial2.Mask = "(##) ####-####"
End Sub

Private Sub mskResidencial2_LostFocus()
   If mskResidencial2.Text = "(__) ____-____" Then
      mskResidencial2.Mask = ""
      mskResidencial2.Text = ""
   End If
End Sub

Private Sub optCOM0800_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyTab Then mskComercial1.SetFocus
End Sub

Private Sub optCOM0800_LostFocus()
   mskComercial1.SetFocus
End Sub

Private Sub txtCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then Exit Sub
   
   If cmdAlterar.Enabled = True Then
      sSQL = "SELECT * FROM telefone WHERE (codigo = " & txtCodigo.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      Campos_Brancos
      frmCadastro.Enabled = True
      
      txtNome.Text = r("nome")
      mskResidencial1.Text = ValidateNull(r("residencial1"))
      mskResidencial2.Text = ValidateNull(r("residencial2"))
      mskComercial1.Text = ValidateNull(r("comercial1"))
      mskComercial2.Text = ValidateNull(r("comercial2"))
      cboCelularOP1.Text = ValidateNull(r("op1"))
      mskCelular1.Text = ValidateNull(r("celular1"))
      cboCelularOP2.Text = ValidateNull(r("op2"))
      mskCelular2.Text = ValidateNull(r("celular2"))
      'If Not IsNull(RS!com0800) Then optCOM0800.Text = RS!com0800
      txtEmail.Text = ValidateNull(r("email"))
      
      On Error Resume Next
      txtNome.SetFocus
   End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   If cmdSalvar.Enabled = False And cmdAlterar.Enabled = False Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(LCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If cmdSalvar.Enabled = False And cmdAlterar.Enabled = False Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub
