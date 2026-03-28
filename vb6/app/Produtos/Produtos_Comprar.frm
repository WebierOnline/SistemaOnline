VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Comprar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUTOS À COMPRAR"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "Produtos_Comprar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   60
      TabIndex        =   10
      Top             =   840
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   11033
      _Version        =   393216
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
      TabCaption(0)   =   "SAÍDA"
      TabPicture(0)   =   "Produtos_Comprar.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExcluir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSair"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmCadastro"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Produtos_Comprar.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid_Historico"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Produtos_Comprar.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdExibir"
      Tab(2).Control(1)=   "Grid_Consulta"
      Tab(2).ControlCount=   2
      Begin VB.Frame frmCadastro 
         Caption         =   "Cadastro"
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
         ForeColor       =   &H00000080&
         Height          =   5655
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   6495
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   17
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   1260
            Width           =   4215
         End
         Begin VB.ComboBox cboProduto 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   6255
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente (Interessado)"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1020
            Width           =   1440
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   555
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   14
         Top             =   420
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   8599
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
         Height          =   4395
         Left            =   -74940
         TabIndex        =   15
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7752
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   6720
         TabIndex        =   7
         Top             =   3780
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Comprar.frx":2426
         PICN            =   "Produtos_Comprar.frx":2442
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
         Left            =   6720
         TabIndex        =   3
         Top             =   1140
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Comprar.frx":275C
         PICN            =   "Produtos_Comprar.frx":2778
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
         Left            =   6720
         TabIndex        =   5
         Top             =   2460
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Comprar.frx":9042
         PICN            =   "Produtos_Comprar.frx":905E
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
         Left            =   6720
         TabIndex        =   4
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Comprar.frx":9938
         PICN            =   "Produtos_Comprar.frx":9954
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
         Left            =   6720
         TabIndex        =   6
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Comprar.frx":103F8
         PICN            =   "Produtos_Comprar.frx":10414
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
         Left            =   6720
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Comprar.frx":1072E
         PICN            =   "Produtos_Comprar.frx":1074A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   735
         Left            =   -67920
         TabIndex        =   18
         Top             =   5400
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "&Exibir"
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
         MICON           =   "Produtos_Comprar.frx":11424
         PICN            =   "Produtos_Comprar.frx":11440
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   8685
      TabIndex        =   8
      Top             =   60
      Width           =   8715
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Produtos_Comprar.frx":11D1A
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTOS À COMPRAR"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   180
         Width           =   3780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   19
      Top             =   7200
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "10:44"
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
Attribute VB_Name = "Produtos_Comprar"
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
   sSQL = "INSERT INTO produtos_comprar (codigo, cod_cliente, produto) VALUES (" & _
      txtCodigo.Text & ", " & IIf(txtCodCliente.Text = "", "Null", txtCodCliente.Text) & ", '" & cboProduto.Text & "');"
   
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
   sSQL = "UPDATE produtos_comprar SET " & _
      "cod_cliente = " & IIf(txtCodCliente.Text = "", "Null", txtCodCliente.Text) & ", " & _
      "produto = " & cboProduto.Text
   
   'Condição para atualização
   sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM produtos_comprar;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodigo.Text = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub FormatarGridConsulta(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Consulta
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 4300
      .ColWidth(4) = 4000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "COD.CLIENTE"
      .TextMatrix(0, 3) = "PRODUTO"
      .TextMatrix(0, 4) = "INTERESSADO"
      .Redraw = False
      
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("var_cod")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_codcliente") & ""
            .TextMatrix(.Rows - 1, 3) = rTabela("var_prod")
            .TextMatrix(.Rows - 1, 4) = rTabela("var_cliente") & ""
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      'For i = 1 To .Rows - 1
      '   .Row = i
      '   .Col = 5
      '   .CellForeColor = &HC0&
      '   .CellFontBold = True
      ' Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Limpar_Objetos()
   txtCodigo.Text = ""
   cboProduto.Text = ""
   txtCodCliente.Text = ""
   CboCliente.Text = ""
End Sub

Private Sub CboCliente_LostFocus()
   On Error GoTo TrataErro
   
   If CboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   If CboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = CboCliente.ItemData(CboCliente.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub CboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   CboCliente.Clear
   
   sSQL = "SELECT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      CboCliente.AddItem r("nome")
      CboCliente.ItemData(CboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo CboCliente
End Sub

Private Sub cboProduto_GotFocus()
   moCombo.AttachTo cboProduto
End Sub

Private Sub cboProduto_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Or cboProduto.Text = "" Or CboCliente.Text = "" Then Exit Sub
   
   'Não é necessário consulta o registro antes de atualiza-lo
   sSQL = "SELECT * FROM produtos_comprar WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   frmCadastro.Enabled = False
   cmdExibir_Click
End Sub

Private Sub cmdCancelar_Click()
   Limpar_Objetos
   frmCadastro.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   If txtCodigo.Text = "" Or cboProduto.Text = "" Then Exit Sub
   
   'Solicita a confirmação do usuário
   If ShowMsg("Excluir esse produto?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'Não é necessário consulta o registro antes de exclui-lo
   'sSQL = "SELECT * FROM produtos_comprar WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)

   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM produtos_comprar WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   frmCadastro.Enabled = False
   cmdExibir_Click
End Sub

Private Sub cmdExibir_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT produtos_comprar.codigo AS var_cod, produtos_comprar.cod_cliente AS var_codcliente, " & _
      "produtos_comprar.produto AS var_prod, cliente.nome AS var_cliente " & _
      "FROM produtos_comprar LEFT JOIN cliente ON produtos_comprar.cod_cliente = cliente.codigo " & _
      "ORDER BY produtos_comprar.produto;"
   
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGridConsulta r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdNovo_Click()
   frmCadastro.Enabled = True
   cmdSalvar.Enabled = True
   cmdCancelar.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   Limpar_Objetos
   Auto_Numeracao
   cboProduto.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   On Error GoTo TrataErro
   
   If txtCodigo.Text = "" Or cboProduto.Text = "" Then
      ShowMsg "Cadastro Incompleto!", vbInformation
      Exit Sub
   End If
   
   'Não é necessário consultar todos os registros antes de inserir um novo
   'sSQL = "SELECT * FROM produtos_comprar;"
   'Set r = BD.OpenRecordset(sSQL)
   
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
   frmCadastro.Enabled = False
   cmdNovo.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdExibir_Click
   Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
   Set moCombo = New cComboHelper
   
   cmdExibir_Click
   Call PreencheProdutos
   
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   
   frmCadastro.Enabled = False
   cmdNovo.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   SSTab1.Tab = 0
End Sub

Sub PreencheProdutos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim var_cboTexto As String
   
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If cboProduto.Text <> "" Then var_cboTexto = cboProduto.Text
   cboProduto.Clear

   Do While Not r.EOF
      cboProduto.AddItem ValidateNull(r("descricao"))
      cboProduto.ItemData(cboProduto.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboProduto.Text = var_cboTexto
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Consulta_DblClick()
   'cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   frmCadastro.Enabled = True
   Limpar_Objetos
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid_Consulta.TextMatrix(Grid_Consulta.Row, 1))
   SSTab1.Tab = 0
End Sub

Private Sub txtCodigo_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

If cmdExcluir.Enabled = True Then
   sSQL = "SELECT * FROM produtos_comprar WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      cboProduto.Text = r("produto")
      txtCodCliente.Text = ValidateNull(r("cod_cliente"))
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If
End Sub

Private Sub TxtCodCliente_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodCliente.Text = "" Then Exit Sub
   
   If cmdExcluir.Enabled = True Then
      sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then CboCliente.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub
