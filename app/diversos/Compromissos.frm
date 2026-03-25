VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Compromissos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COMPROMISSOS"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "Compromissos.frx":0000
   LinkTopic       =   "Form73"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Mostrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   17
      Top             =   5100
      Width           =   4335
      Begin MSMask.MaskEdBox MskDataNova 
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton optFeito 
         Caption         =   "Feito"
         Height          =   195
         Left            =   1020
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton optFazer 
         Caption         =   "À fazer"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   795
      End
      Begin VB.CommandButton Command11 
         Caption         =   ">>"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "<<"
         Height          =   255
         Left            =   3420
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "ok"
         Height          =   255
         Left            =   2940
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro de Compromissos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      TabIndex        =   11
      Top             =   5100
      Width           =   7155
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6720
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Compromissos.frx":23D2
         Left            =   120
         List            =   "Compromissos.frx":23D4
         TabIndex        =   0
         Top             =   480
         Width           =   1155
      End
      Begin VB.ComboBox cboNome 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2835
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   5940
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   4035
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Para"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   900
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   5940
         TabIndex        =   16
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tarefa:"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   2520
         TabIndex        =   14
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Encarregado"
         Height          =   195
         Left            =   3600
         TabIndex        =   12
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5220
      Visible         =   0   'False
      Width           =   1080
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   555
      Left            =   8820
      TabIndex        =   25
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Compromissos.frx":23D6
      PICN            =   "Compromissos.frx":23F2
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
      Height          =   555
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Compromissos.frx":270C
      PICN            =   "Compromissos.frx":2728
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
      Height          =   555
      Left            =   1860
      TabIndex        =   27
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Compromissos.frx":8FF2
      PICN            =   "Compromissos.frx":900E
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
      Height          =   555
      Left            =   3600
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Compromissos.frx":FAB2
      PICN            =   "Compromissos.frx":FACE
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
      Height          =   555
      Left            =   5340
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Compromissos.frx":103A8
      PICN            =   "Compromissos.frx":103C4
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
      Height          =   555
      Left            =   7080
      TabIndex        =   30
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Compromissos.frx":106DE
      PICN            =   "Compromissos.frx":106FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "hora   encarregado   tipo   para    tarefa"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   420
      Width           =   9735
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROMISSOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Width           =   11580
   End
End
Attribute VB_Name = "Compromissos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Private Function Atualizar_Dados() As Boolean
   
   'TBcompromissos("CODIGO") = txtCodigo.Text
   'TBcompromissos("DATA") = MaskEdBox1.Text
   'TBcompromissos("HORA") = MaskEdBox2.Text
   'TBcompromissos("COMPROMISSO") = Text1.Text
   'TBcompromissos("ENCARREGADO") = Combo1.Text
   'TBcompromissos("TIPO") = cboTipo.Text
   'TBcompromissos("STATUS") = cboStatus.Text
   'TBcompromissos("PARA") = cboNome.Text
   
End Function

Private Sub Atualizar_Dados2()
   'dYn("CODIGO") = txtCodigo.Text
   'dYn("DATA") = MaskEdBox1.Text
   'dYn("HORA") = MaskEdBox2.Text
   'dYn("COMPROMISSO") = Text1.Text
   'dYn("ENCARREGADO") = Combo1.Text
   'dYn("TIPO") = cboTipo.Text
   'dYn("STATUS") = cboStatus.Text
   'dYn("PARA") = cboNome.Text
End Sub

Private Sub LIMPAR_DADOS()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   Text1.Text = ""
   Combo1.Text = ""
   cboTipo.Text = ""
   txtCodigo.Text = ""
   cboStatus.Text = ""
   cboNome.Text = ""
End Sub

Private Sub Mostrar_Dados()
   'txtCodigo.Text = TBcompromissos("CODIGO")
   'MaskEdBox1.Text = TBcompromissos("DATA")
   'MaskEdBox2.Text = TBcompromissos("HORA")
   'Text1.Text = TBcompromissos("COMPROMISSO")
   'Combo1.Text = TBcompromissos("ENCARREGADO")
   'cboTipo.Text = TBcompromissos("TIPO")
   'cboNome.Text = TBcompromissos("PARA")
   'cboStatus.Text = TBcompromissos("STATUS")
End Sub

Private Sub cboNome_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboNome.Clear
   
   sSQL = "SELECT nome FROM funcionario ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboNome.AddItem r("nome")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboNome
End Sub

Private Sub cboStatus_GotFocus()
   cboStatus.Clear
   cboStatus.AddItem "À fazer"
   cboStatus.AddItem "Feito"
   cboStatus.ListIndex = 0
   moCombo.AttachTo cboStatus
End Sub

Private Sub cboTipo_GotFocus()
   cboTipo.Clear
   cboTipo.AddItem "Fazer"
   cboTipo.AddItem "Ir"
   cboTipo.AddItem "Ligar"
   cboTipo.AddItem "Entregar"
   cboTipo.AddItem "Pedir"
   moCombo.AttachTo cboTipo
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If Combo1.Text = "" Or Text1.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   sSQL = "SELECT * FROM compromissos WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      Atualizar_Dados2
      
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   ShowMsg "Operação concretizada com sucesso", vbInformation
   LIMPAR_DADOS
   Form_Load
   cboStatus.SetFocus
   Exit Sub
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If Combo1.Text = "" Or Text1.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   If MsgBox("Tem certeza que deseja excluir este registro?", vbInformation + vbYesNo, "Aviso do Sistema") = vbNo Then
      cboStatus.SetFocus
      Exit Sub
   End If
   
   sSQL = "SELECT * FROM compromisso WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      
      ShowMsg "Operação realizada com sucesso.", vbExclamation
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   LIMPAR_DADOS
   Form_Load
   cboStatus.SetFocus
End Sub

Private Sub cmdNovo_Click()
LIMPAR_DADOS
   Form_Load
   cboStatus.SetFocus
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   
   If Text1.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
      Exit Sub
   Else
      If Not Atualizar_Dados Then
         
         Exit Sub
      End If
      
      LIMPAR_DADOS
      Form_Load
   End If
   
   cboStatus.SetFocus
   Exit Sub
End Sub

Private Sub Combo1_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Combo1.Clear
   
   sSQL = "SELECT nome FROM funcionario;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      Combo1.AddItem r("nome")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo Combo1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command10_Click()
   'If TBcompromissos.BOF Then Exit Sub
   'TBcompromissos.MovePrevious
   'If TBcompromissos.BOF Then Exit Sub
   'Mostrar_Dados
End Sub

Private Sub Command11_Click()
   'If TBcompromissos.EOF Then Exit Sub
   'TBcompromissos.MoveNext
   'If TBcompromissos.EOF Then Exit Sub
   'Mostrar_Dados
End Sub

Private Sub Command12_Click()
   If MskDataNova.Text = "" Then Exit Sub
   If MskDataNova.Text = "__/__/__" Then Exit Sub
   
   If optFazer.Value = True Then
   '   ABRIR_BD_com_Data Me.Data1
   '   Data1.RecordSource = "SELECT * FROM COMPROMISSOS WHERE DATA = #" & Format(MskDataNova, "mm/dd/yy") & "# AND STATUS = 'À fazer' order by hora, TIPO"
   '   Data1.Refresh
   
   ElseIf optFeito.Value = True Then
   '   ABRIR_BD_com_Data Me.Data1
   '   Data1.RecordSource = "SELECT * FROM COMPROMISSOS WHERE DATA = #" & Format(MskDataNova, "mm/dd/yy") & "# AND STATUS = 'Feito' order by hora, TIPO"
   '   Data1.Refresh
   End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command6_Click()
   Unload Me
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()

End Sub

Private Sub DBGrid1_DblClick()
   'DBGrid1.Col = 0
   'MaskEdBox2.Text = DBGrid1.Text
   'DBGrid1.Col = 1
   'Combo1.Text = DBGrid1.Text
   'DBGrid1.Col = 2
   'cboTipo.Text = DBGrid1.Text
   'DBGrid1.Col = 3
   'cboNome.Text = DBGrid1.Text
   'DBGrid1.Col = 4
   'txtCodigo.Text = DBGrid1.Text
   'DBGrid1.Col = 5
   'cboStatus.Text = DBGrid1.Text
   'DBGrid1.Col = 6
   'MaskEdBox1.Text = DBGrid1.Text
   'DBGrid1.Col = 7
   'Text1.Text = DBGrid1.Text
   'cboStatus.SetFocus
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Data1.Recordset.RecordCount Then
      'DBGrid1.SelBookmarks.Add Data1.Recordset.Bookmark
   End If
End Sub

Private Sub Form_Load()
   Dim i As Long
   
   'Set TBcompromissos = BD.OpenRecordset("COMPROMISSOS", dbOpenTable)
   'TBcompromissos.Index = ("index_codigo")
   '
   'If TBcompromissos.RecordCount = 0 Then
   '   txtCodigo.Text = "1"
   'Else
   '   TBcompromissos.MoveLast
   '   i = TBcompromissos!Codigo + 1
   '   txtCodigo.Text = i
   'End If
   
   MaskEdBox1 = Format(Date, "dd/mm/yy")
   
   'ABRIR_BD_com_Data Me.Data1
   'Data1.RecordSource = "SELECT * FROM COMPROMISSOS WHERE STATUS = 'À fazer' order by DATA, HORA"
   'Data1.Refresh
   
   Set moCombo = New cComboHelper
End Sub

Function Maiuscula(KeyAscii As Integer)
   If KeyAscii > 96 And KeyAscii < 123 Then
      KeyAscii = KeyAscii - 32
   End If
   Maiuscula = KeyAscii
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
   'TBcompromissos.Close
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
   MaskEdBox1.Mask = "##/##/##"
End Sub

Private Sub MaskEdBox1_LostFocus()
   If MskDataNova.Text = "" Then
      Exit Sub
   ElseIf MskDataNova.Text = "__/__/__" Then
      Exit Sub
   ElseIf Not IsDate(MaskEdBox1) Then
      ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
      MaskEdBox1.SetFocus
      MaskEdBox1.SelStart = 0
      MaskEdBox1.SelLength = Len(MaskEdBox1)
   End If
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
   MaskEdBox2.Mask = "##:##"
End Sub

Private Sub MskDataNova_LostFocus()
   If MskDataNova.Text = "__/__/__" Then
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   
   If Not IsDate(MskDataNova) Then
      ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
      MskDataNova.SetFocus
      MskDataNova.SelStart = 0
      MskDataNova.SelLength = Len(MskDataNova)
   Exit Sub
   End If
End Sub

Private Sub optFazer_Click()
   MskDataNova.SetFocus
End Sub

Private Sub optFeito_Click()
   MskDataNova.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Maiuscula(KeyAscii)
End Sub
