VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form BuscaGrid 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   9015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   9015
      TabIndex        =   1
      Top             =   60
      Width           =   9015
      Begin VB.OptionButton optDesc 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descrição"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optRef 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Referência"
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   60
         Width           =   1275
      End
      Begin VB.OptionButton optFab 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fabricante"
         Height          =   195
         Left            =   2820
         TabIndex        =   3
         Top             =   60
         Width           =   1275
      End
      Begin VB.OptionButton optTam 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tamanho"
         Height          =   195
         Left            =   4140
         TabIndex        =   2
         Top             =   60
         Width           =   1275
      End
   End
   Begin MSComctlLib.ListView lstBusca 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "BuscaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCancelled As Boolean
Dim vInfo() As String

Public Property Get Cancelled() As Boolean
   Cancelled = mCancelled
End Property

Public Property Get InfoProduct() As String()
   InfoProduct = vInfo
End Property

Sub pCriarGrid()
   lstBusca.FullRowSelect = True
   lstBusca.LabelEdit = lvwManual
   lstBusca.Visible = True
   lstBusca.View = lvwReport
   lstBusca.HideSelection = False
   lstBusca.ListItems.Clear
   
   lstBusca.ColumnHeaders.Clear
   lstBusca.ColumnHeaders.Add , , "CÓDIGO", 0
   lstBusca.ColumnHeaders.Add , , "COD_BARRA", 0
   lstBusca.ColumnHeaders.Add , , "DESCRIÇÃO", 3200
   lstBusca.ColumnHeaders.Add , , "REF.", 1200
   lstBusca.ColumnHeaders.Add , , "TAM.", 900
   lstBusca.ColumnHeaders.Add , , "FABRICANTE", 1600
   lstBusca.ColumnHeaders.Add , , "QTDE", 1050, 1
   lstBusca.ColumnHeaders.Add , , "VALOR", 800, 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 0 Then
      If KeyCode = vbKeyEscape Then Unload Me
'      If KeyCode = vbKeyReturn Then lstBusca_KeyDown KeyCode, Shift
   End If
End Sub

Private Sub Form_Load()
'Não vender produtos zerados
'Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
'bEstNeg = CBool(oCfg.Value)
'Set oCfg = Nothing

Set Icon = Nothing
KeyPreview = True
mCancelled = True
Erase vInfo
pCriarGrid
End Sub

Private Sub lstBusca_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim vQtde As Double
   'Verifica_QuantEstoque
   
   If Shift = 0 And KeyCode = vbKeyReturn Then
      If lstBusca.ListItems.Count = 0 Then
         ShowMsg "Nenhum item disponível para seleção.", vbExclamation
         Exit Sub
      End If
      
      If lstBusca.SelectedItem Is Nothing Then
         ShowMsg "Nenhum item foi selecionado.", vbExclamation
         Exit Sub
      End If
      
      If Not lstBusca.SelectedItem.Selected Then
         ShowMsg "Nenhum item foi selecionado.", vbExclamation
         Exit Sub
      End If
      
      'If lstBusca.SelectedItem.SubItems(6) = "" Then
      '   ShowMsg "Não há quantidade em estoque informada.", vbExclamation
      '   Exit Sub
      'End If
      
      'Calcula o saldo atual em estoque
      vQtde = EstoqueVendas(lstBusca.SelectedItem.Text)
      
      If vQtde <= 0 Then
         Dim oCfg As ConfigItem
         Dim bEstNeg As Boolean
         
         'Recupera a configuração do estoque
         Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
         bEstNeg = CBool(oCfg.Value)
         Set oCfg = Nothing
         
         If Not bEstNeg Then
            ShowMsg "A quantidade em estoque é insuficiente.", vbExclamation
            Exit Sub
         End If
      End If
      
      
      ReDim vInfo(1 To 3)
      vInfo(1) = lstBusca.SelectedItem
      vInfo(2) = lstBusca.SelectedItem.ListSubItems.Item(1).Text
      vInfo(3) = lstBusca.SelectedItem.ListSubItems.Item(2).Text
      
      mCancelled = False
      Unload Me
   End If

End Sub

Private Sub optFab_Click()
   'txtDescricao.SetFocus
End Sub

Private Sub optRef_Click()
   'txtDescricao.SetFocus
End Sub

Private Sub optTam_Click()
   'txtDescricao.SetFocus
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   On Local Error Resume Next
   If KeyAscii = 13 Then
      'If lstBusca.Visible = True And Not IsNumeric(txtCodBarra.Text) Then
      '   lstBusca.SetFocus
      'Else
      SendKey ocKEYTAB
         'Mostrar_Descricao_Produto
         'Adicionar_Produto
         'MostrarGrid_Produtos
         
      'End If
   End If
End Sub

Private Sub txtDescricao_Validate(Cancel As Boolean)
   Dim sSQL As String
   Dim fSQL As String
   Dim r As ADODB.Recordset
   Dim ItemLst As ListItem
   Dim fTam As String

Dim varSeVendeNegativo As String
If bEstNeg = False Then
    varSeVendeNegativo = " AND (produtos.quant_estoque > 0)"
Else
    varSeVendeNegativo = " "
End If
   
   
   
   If optRef.Value = True Then
      fSQL = "(ref LIKE '%" & txtDescricao.Text & "%')"
   
   ElseIf optFab.Value = True Then
      fSQL = "(fabricante LIKE '%" & txtDescricao.Text & "%')"
   
   ElseIf optTam.Value = True Then
      fTam = MontarCriterios(txtDescricao)
      
      If fTam = "#1" Then
         ShowMsg "Intervalo de tamanho incompatível.", vbExclamation
         Exit Sub
      End If
            
      fSQL = "(tamanho IN (" & fTam & "))"
   
   ElseIf optDesc.Value = True Then
      fSQL = "(descricao LIKE '%" & txtDescricao.Text & "%')"
   
   End If
   
   sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, " & _
          "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
          "FROM produtos WHERE " & fSQL & " AND (produtos.ativo = 1)  " & varSeVendeNegativo & " " & _
          "ORDER BY descricao;"
      
   lstBusca.ListItems.Clear
   lstBusca.Refresh
   
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      'primeira coluna
      Set ItemLst = lstBusca.ListItems.Add(, , r("var_cod"))
      'segunda e terceira coluna, que são sub itens da coluna 1
      ItemLst.SubItems(1) = ValidateNull(r("var_codbarra"))
      ItemLst.SubItems(2) = ValidateNull(r("var_desc"))
      ItemLst.SubItems(3) = ValidateNull(r("var_ref"))
      ItemLst.SubItems(4) = ValidateNull(r("var_tam"))
      ItemLst.SubItems(5) = ValidateNull(r("var_fab"))
      If Not IsNull(r("var_quant")) Then ItemLst.SubItems(6) = r("var_quant")
      If Not IsNull(r("venda")) Then ItemLst.SubItems(7) = Format(r("venda"), ocMONEY)
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Verifica_QuantEstoque()
   Dim VERIFICAR_QUANTIDADE As Boolean
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim oCfg As ConfigItem
   Dim bEstNeg As Boolean
   
   'Recupera a configuração do estoque
   Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
   bEstNeg = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   If Not bEstNeg Then
      sSQL = "SELECT codigo, quant_estoque FROM produtos WHERE 0 = 1"
      '      sSQL = "SELECT codigo, quant_estoque FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "');"

      Set r = dbData.OpenRecordset(sSQL)
      
      VERIFICAR_QUANTIDADE = False
      'If txtQuant.Text = "" Then txtQuant.Text = 0
      If r.EOF Then Exit Sub
      
      If r("quant_estoque") <= 0 Then
         ShowMsg "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE 0 (zero) PRODUTO", vbExclamation
         'LimparObjetos_Pedido
         'cmdAlterar.Enabled = False
         VERIFICAR_QUANTIDADE = True
         'txtCodBarra.Text = ""
      End If
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   Else
      Exit Sub
   End If
End Sub
