VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Produtos_CadastoRapido 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro Rápido"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuant 
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Top             =   300
      Width           =   855
   End
   Begin VB.TextBox txtCusto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7860
      TabIndex        =   4
      Top             =   300
      Width           =   1155
   End
   Begin VB.TextBox txtCodBarra 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      MaxLength       =   90
      TabIndex        =   0
      Top             =   300
      Width           =   1935
   End
   Begin VB.ComboBox cboUnidMedida 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   6900
      TabIndex        =   3
      Top             =   300
      Width           =   915
   End
   Begin VB.TextBox txtDescricao 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2040
      MaxLength       =   90
      TabIndex        =   1
      Top             =   300
      Width           =   3915
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   555
      Left            =   7200
      TabIndex        =   6
      Top             =   660
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
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
      MICON           =   "Produtos_CadastoRapido.frx":0000
      PICN            =   "Produtos_CadastoRapido.frx":001C
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
      Left            =   5340
      TabIndex        =   5
      Top             =   660
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
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
      MICON           =   "Produtos_CadastoRapido.frx":1DAE
      PICN            =   "Produtos_CadastoRapido.frx":1DCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQuantAtual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant."
      Height          =   195
      Left            =   6000
      TabIndex        =   11
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda"
      Height          =   195
      Left            =   7860
      TabIndex        =   10
      Top             =   60
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Barra"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   750
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unid. Med."
      Height          =   195
      Left            =   6900
      TabIndex        =   8
      Top             =   60
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "Produtos_CadastoRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim varCodProduto As Integer
Dim vQuant As Double
Private Sub cboUnidMedida_GotFocus()
Dim var_Texto As String
var_Texto = cboUnidMedida.Text

   cboUnidMedida.Clear
   cboUnidMedida.AddItem "UN"
   cboUnidMedida.AddItem "CX"
   cboUnidMedida.AddItem "M"
   cboUnidMedida.AddItem "M²"
   cboUnidMedida.AddItem "M³"
   cboUnidMedida.AddItem "ML"
   cboUnidMedida.AddItem "KG"
   cboUnidMedida.AddItem "GR"
   moCombo.AttachTo cboUnidMedida
   
cboUnidMedida.Text = var_Texto
cboUnidMedida.SelStart = 0
cboUnidMedida.SelLength = Len(cboUnidMedida)
End Sub




Private Sub cmdCancelar_Click()
txtCodBarra.Text = ""
txtDescricao.Text = ""
txtQuant.Text = ""
cboUnidMedida.Text = ""
txtCusto.Text = ""
Unload Me
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
varCodProduto = 0
cboUnidMedida.Text = "UN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub txtCodBarra_LostFocus()
If txtCodBarra.Text = "" Then
    Dim sSQL As String
    Dim r As ADODB.Recordset
    sSQL = "SELECT isnull(MAX(COD_BARRA), 0) as UltimoCodigo FROM produtos where len(COD_BARRA) = 5;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Val(Len(txtCodBarra)) < 13 Then
        Dim vCodBarraInt As String
        vCodBarraInt = Val(r("UltimoCodigo"))
    End If
    
    If Not r.BOF Then
        txtCodBarra.Text = Format(vCodBarraInt + 1, "00000")
    End If
Else
    If Len(txtCodBarra) < 13 And txtCodBarra.Text <> "" Then
        txtCodBarra.Text = Format(txtCodBarra.Text, "00000")
    ElseIf Len(txtCodBarra) > 13 Then
        MsgBox "Esse Cód. de Barra possui mais números que o permitido", vbInformation, "Aviso do Sistema"
        txtCodBarra.SetFocus
        Exit Sub
    End If
End If
End Sub




Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Function TirarEspaco(ByVal Value As String) As String
Dim bRepete As Boolean
Value = Replace$(Value, "'", vbNullString)
Do
  Value = Replace$(Value, "  ", " ")
  bRepete = InStr(1, Value, "  ", vbTextCompare)
  Value = Trim(Value)
Loop Until Not bRepete

TirarEspaco = Value
End Function
Private Sub txtCodBarra_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" Then Exit Sub
txtCodBarra.Text = Trim(txtCodBarra.Text)

'Verifica se existe o código de barras cadastrado
sSQL = "SELECT codigo, ativo, cod_barra FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

'If cmdAlterar.Enabled = False Then
   If r.RecordCount > 0 Then
        If r("ativo") = True Then
            ShowMsg "Já existe um produto cadastrado com esse cód. de barra!", vbInformation
            Cancel = True           'Cancela a entrada e permanece com o foco no campo
            txtCodBarra.Text = ""   'Limpa a entrada
            txtCodBarra.SetFocus
            Exit Sub                'Evita a saída do campo
        ElseIf r("ativo") = False Then
            ShowMsg "Existe um produto DESABILITADO com esse cód. de barra!", vbInformation
            Cancel = True           'Cancela a entrada e permanece com o foco no campo
            txtCodBarra.Text = ""   'Limpa a entrada
            txtCodBarra.SetFocus
            Exit Sub
        End If
   End If
'End If
End Sub

Private Sub txtCusto_GotFocus()
txtCusto.SelStart = 0
txtCusto.SelLength = Len(txtCusto)
End Sub

Private Sub txtCusto_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub



Private Sub txtCusto_LostFocus()
Dim varLucro As Currency

If txtCusto.Text = "" Then Exit Sub
varLucro = txtCusto.Text

txtCusto.Text = FormatNumber(varLucro, 2)

'CalcularPrecos
End Sub



Private Sub cmdSalvar_Click()
If txtDescricao.Text = "" Then ShowMsg "Digite a Descrição do produto", vbInformation: txtDescricao.SetFocus: Exit Sub
If txtCusto.Text = "" Then ShowMsg "Produtos estão sem margens de vendas", vbInformation: txtCusto.SetFocus: Exit Sub
If txtCodBarra.Text = "" Then MsgBox "Não será permitido cadastrar produto sem código de barra", vbInformation, "Aviso do Sistema": Exit Sub

If txtQuant.Text = "" Then
    vQuant = 0
Else
    vQuant = txtQuant.Text
End If

AutoNumeracao

If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

Preco_Entrada
Quant_Entrada

txtCodBarra.Text = ""
txtDescricao.Text = ""
txtQuant.Text = ""
cboUnidMedida.Text = ""
txtCusto.Text = ""
Unload Me
End Sub

Private Sub AutoNumeracao()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_produto FROM produtos;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then varCodProduto = r("cod_produto") + 1
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Function Inserir_Dados() As Boolean
Dim sSQL As String

'Comando de inclusão
sSQL = "INSERT INTO produtos (" & _
   "codigo, destaque, cod_barra, ean, descricao, fabricante, unid_medida, " & _
   "categoria, PRATELEIRA, quant_min, observacao, quant_estoque, ref, tamanho, ICMSCST, ICMSAliq, PISCST, COFINSCST, IPICST, NCM, CEST, CFOP, Alterado, ativo, PEDIRPESO, PISAliq, COFINSAliq, IPIAliq, USOCONSUMO, COMBUSTIVEL, MATERIAPRIMA, IMOBILIZADO, FRACIONADO) VALUES (" & _
   varCodProduto & ", 0, '" & txtCodBarra.Text & "', '" & txtCodBarra.Text & "', '" & txtDescricao.Text & "', '', '" & cboUnidMedida.Text & "', '', '', 0, '', " & Replace(CDbl(vQuant), ",", ".") & ", '', '', '0', 0, '08', '08', '0', '0', '0', '0', 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0);"

   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function


Private Sub Preco_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
'If cmdSalvar.Enabled = True Then
   Dim AutoNumeracao As Long
   
   'AUTONUMERAÇÃO
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_precos;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

 
   sSQL = "INSERT INTO produtos_precos (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP, CUSTO) VALUES (" & _
      AutoNumeracao & ", " & varCodProduto & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 0, 'CADASTRO', 0, " & Replace(CCur(txtCusto.Text), ",", ".") & ", 0, 0, 0, 0, 0, 0, 0 );"
   dbData.Execute sSQL
'End If
End Sub
Private Sub Quant_Entrada()
Dim sSQL As String
Dim r As ADODB.Recordset

'ENTRADA DO PRODUTO
'If cmdSalvar.Enabled = True Then
   Dim AutoNumeracao As Long
   
   'AUTONUMERAÇÃO
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_quant;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

   
   sSQL = "INSERT INTO produtos_quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO) VALUES (" & _
      AutoNumeracao & ", " & varCodProduto & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), 0, 'CADASTRO', " & Replace(CDbl(vQuant), ",", ".") & ", 'ADIÇÃO');"
   dbData.Execute sSQL
'End If
End Sub


Private Sub txtDescricao_LostFocus()
txtDescricao.Text = TirarEspaco(txtDescricao.Text)
End Sub


