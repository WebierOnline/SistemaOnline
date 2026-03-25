VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_AdicionarQuant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADICIONAR QUANTIDADE"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   Icon            =   "Produtos_AdicionarQuant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Alteração"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   60
      TabIndex        =   14
      Top             =   2040
      Width           =   11415
      Begin VB.TextBox txtQuantNova 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin ChamaleonBtn.chameleonButton cmdAdicionar 
         Height          =   315
         Left            =   2340
         TabIndex        =   17
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Adicionar"
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
         MICON           =   "Produtos_AdicionarQuant.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdRemover 
         Height          =   315
         Left            =   4020
         TabIndex        =   18
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Remover"
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
         MICON           =   "Produtos_AdicionarQuant.frx":23EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
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
      TabIndex        =   4
      Top             =   960
      Width           =   11415
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtCodProduto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtFabricante 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtQuant 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód."
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante"
         Height          =   195
         Left            =   7980
         TabIndex        =   10
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant. Atual"
         Height          =   195
         Left            =   10020
         TabIndex        =   9
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   11385
      TabIndex        =   1
      Top             =   120
      Width           =   11415
      Begin VB.TextBox txtCodUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7440
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Produtos_AdicionarQuant.frx":240A
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ADICIONAR QUANTIDADE"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   180
         Width           =   3870
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   3735
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16060
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "19:14"
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
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   615
      Left            =   9300
      TabIndex        =   13
      Top             =   3060
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
      MICON           =   "Produtos_AdicionarQuant.frx":877A
      PICN            =   "Produtos_AdicionarQuant.frx":8796
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
Attribute VB_Name = "Produtos_AdicionarQuant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalvar_Click()
End Sub
Private Sub MostrarDados_Produto(rTabela As ADODB.Recordset)
txtDescricao.Text = ValidateNull(rTabela("descricao"))
txtFabricante.Text = ValidateNull(rTabela("fabricante"))
txtQuant.Text = ValidateNull(rTabela("quant_estoque"))
End Sub

Private Sub cmdAdicionar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim AutoNumeracao As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM Produtos_Quant;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

If ShowMsg("Deseja adicionar mais itens no produto: " & txtDescricao.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

sSQL = "INSERT INTO Produtos_Quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _
   AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 0, 'AJUSTE', " & Replace(CDbl(txtQuantNova.Text), ",", ".") & ", 'ADIÇÃO', '" & Format(Now, ocHRMN) & "', " & txtCodUsuario.Text & ", " & Replace(CDbl(txtQuant.Text), ",", ".") & ");"
dbData.Execute sSQL

'Atualiza o estoque do produto
dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(txtQuantNova.Text, ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"

cmdSair_Click
End Sub

Private Sub Text1_Change()

End Sub

Private Sub cmdRemover_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim AutoNumeracao As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM Produtos_Quant;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

If ShowMsg("Deseja remover mais itens no produto: " & txtDescricao.Text & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub


sSQL = "INSERT INTO Produtos_Quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _
   AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 0, 'AJUSTE', " & Replace(CDbl(txtQuantNova.Text), ",", ".") & ", 'REMOÇÃO', '" & Format(Now, ocHRMN) & "', " & txtCodUsuario.Text & ", " & Replace(CDbl(txtQuant.Text), ",", ".") & ");"
dbData.Execute sSQL

'Atualiza o estoque do produto
dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(txtQuantNova.Text, ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"

cmdSair_Click
End Sub

Private Sub cmdSair_Click()
LimparObjetos_Produtos
txtCodProduto.Text = ""
Me.Hide
Produtos_Estoque_Simples.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Produtos_Estoque_Simples.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
txtCodProduto.Text = ""
End Sub

Private Sub txtCodProduto_Change()
Dim sSQL As String
Dim r As ADODB.Recordset
If txtCodProduto.Text = "" Then Exit Sub

sSQL = "SELECT * FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

LimparObjetos_Produtos
MostrarDados_Produto r
End Sub
Private Sub LimparObjetos_Produtos()
txtDescricao.Text = ""
txtFabricante.Text = ""
txtQuant.Text = ""
txtQuantNova.Text = ""
End Sub

