VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Acesso [By Tecl@]"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7800
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   50
      ScaleHeight     =   6435
      ScaleWidth      =   8325
      TabIndex        =   4
      Top             =   60
      Width           =   8385
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
         Left            =   6000
         TabIndex        =   7
         ToolTipText     =   "Emitir relatório de usuários cadastrados"
         Top             =   1200
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
         Left            =   120
         TabIndex        =   2
         Top             =   1080
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
         Left            =   6000
         TabIndex        =   1
         ToolTipText     =   "Gravar perfil de acesso do usuário"
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   4455
         Left            =   75
         TabIndex        =   3
         Top             =   1920
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   7858
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
      Begin VB.Image Image1 
         Height          =   450
         Left            =   5880
         Picture         =   "Form1.frx":1172
         ToolTipText     =   "Logo VBMania"
         Top             =   50
         Width           =   2250
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
         TabIndex        =   6
         Top             =   1680
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Combo1_Click()
'---------------------------------------
'Mostrar o perfil do usuário selecionado
'---------------------------------------
rsUsuario.MoveFirst
rsUsuario.Find "login='" & Combo1 & "'"
Check1.Value = 0
If Not rsUsuario.EOF Then
    Call LerAcesso
Else
    MsgBox "Ocorreu um erro ao localizar usuário!", vbExclamation, "Acesso"
End If
End Sub

Private Sub Command1_Click()
'------------------------
'Gravar perfil do usuário
'------------------------
If MsgBox("Confirma a alteração no perfil de acesso?", vbQuestion + vbYesNo + vbDefaultButton2, "Controle de Acesso") = vbYes Then
    rsUsuario.MoveFirst
    rsUsuario.Find "login='" & Combo1 & "'"
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
    Set Imagem.Picture = LoadPicture(App.Path & "\vbmania_logo.jpg")
    Set DataReport1.DataSource = rsUsuario
    DataReport1.Show
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Desconectar
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Form2.Show vbModal
End Sub

Public Sub ListarAcesso()
'--------------------------
'Listar os tipos de acessos
'--------------------------
Do While Not rsAcesso.EOF
    ListView1.ListItems.Add , , rsAcesso.Fields(1)
    rsAcesso.MoveNext
Loop
End Sub

Public Sub ListarUsuario()
'-------------------------
'Listar o Login do usuário
'-------------------------
Do While Not rsUsuario.EOF
    Combo1.AddItem rsUsuario.Fields("login")
    rsUsuario.MoveNext
Loop
Combo1.ListIndex = 0
End Sub

Sub GravarAcesso()
'----------------------------------------------
'Gravar perfil de acesso de usuário selecionado
'----------------------------------------------
For i = 1 To ListView1.ListItems.Count
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
