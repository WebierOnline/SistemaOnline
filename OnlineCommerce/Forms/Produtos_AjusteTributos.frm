VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_AjusteTributos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AJUSTE DE TRIBUTOS"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   Icon            =   "Produtos_AjusteTributos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Produtos_AjusteTributos.frx":1D82
   ScaleHeight     =   9825
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFiltro 
      Caption         =   "Quantidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   29
      Top             =   8520
      Width           =   1695
      Begin VB.OptionButton optMostrarQuant 
         Caption         =   "Com quantidade"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton optMostrarNegativos 
         Caption         =   "Negativos"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton optMostrarZerados 
         Caption         =   "Zerados"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optMostrarTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Preço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   25
      Top             =   8520
      Width           =   1335
      Begin VB.OptionButton optSemPreco 
         Caption         =   "Sem Preço"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optComPreco 
         Caption         =   "Com Preço"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optTodosPreco 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   540
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critérios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   60
      TabIndex        =   13
      Top             =   8520
      Width           =   1395
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   1035
      End
      Begin VB.OptionButton optCodBarra 
         Caption         =   "Cód. Barra"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busca Avançada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6180
      TabIndex        =   8
      Top             =   8520
      Width           =   6645
      Begin VB.TextBox txtCodBarra 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CheckBox chkDescPorIniciais 
         Caption         =   "Por Iniciais"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkDescPorProduto 
         Caption         =   "Por Produto"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cboConsLinha 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.ComboBox cboDesc 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   495
         Left            =   5040
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         MICON           =   "Produtos_AjusteTributos.frx":264C
         PICN            =   "Produtos_AjusteTributos.frx":2668
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodBarra 
         Caption         =   "Cód. de Barra"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   4800
      TabIndex        =   5
      Top             =   8520
      Width           =   1335
      Begin VB.CheckBox ckkORDDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox ckkORDLinha 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Produtos_AjusteTributos.frx":2F42
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   3540
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   15165
      TabIndex        =   0
      Top             =   60
      Width           =   15195
      Begin VB.Image Image1 
         Height          =   645
         Left            =   600
         Picture         =   "Produtos_AjusteTributos.frx":3F7A
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AJUSTE DE TRIBUTOS"
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
         Left            =   1365
         TabIndex        =   1
         Top             =   240
         Width           =   3420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   23
      Top             =   9555
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22701
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "13:30"
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
      Height          =   6975
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   315
      Left            =   13500
      TabIndex        =   24
      Top             =   8040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
      MICON           =   "Produtos_AjusteTributos.frx":A2EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultarNCM 
      Height          =   315
      Left            =   8340
      TabIndex        =   36
      Top             =   8040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consultar NCM pela Descrição"
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
      MICON           =   "Produtos_AjusteTributos.frx":A306
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultaNCMean 
      Height          =   315
      Left            =   10920
      TabIndex        =   37
      Top             =   8040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consultar NCM pelo EAN"
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
      MICON           =   "Produtos_AjusteTributos.frx":A322
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQuant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Left            =   660
      TabIndex        =   35
      Top             =   8040
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant.:"
      Height          =   195
      Left            =   60
      TabIndex        =   34
      Top             =   8040
      Width           =   525
   End
End
Attribute VB_Name = "Produtos_AjusteTributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private moCombo As cComboHelper
Private iRow As Long, iCol As Long
'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Private Sub LimparGrid2()
Dim sSQL As String
Dim r As ADODB.Recordset
   
sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
   "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.UNID_MEDIDA AS var_UnidMed, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE 1 = 0"

Set r = dbData.OpenRecordset(sSQL)

Formatar_Grid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub LimparGrid()
   Dim i As Integer
   
   txtEdit.Text = ""
   
   With Grid
      .Clear
      .Cols = 9
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 4200
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1000
      .ColWidth(8) = 2000
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "QUANT."
      .TextMatrix(0, 6) = "MIN."
      .TextMatrix(0, 7) = "VENDA"
      .TextMatrix(0, 8) = "CATEGORIA"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      'ALINHAMENTO
      .ColAlignment(2) = 1
      
      .Rows = .Rows + 1
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub chkDescPorIniciais_Click()
   If optDesc.Value = Unchecked Then Exit Sub
   
   If chkDescPorIniciais.Value = Checked Then
      cboDesc.Clear
      chkDescPorProduto.Value = Unchecked
      cboDesc.SetFocus
   End If
End Sub

Private Sub chkDescPorProduto_Click()
   If optDesc.Value = Unchecked Then Exit Sub
   
   If chkDescPorProduto.Value = Checked Then
      chkDescPorIniciais.Value = Unchecked
      cboDesc.SetFocus
   End If
End Sub

Private Sub MostrarCriterios()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim var_Criterio As String
   Dim var_Indice As String
   
   var_Criterio = ""
   
   If chkDescPorProduto.Value = Checked Then
      var_Criterio = var_Criterio & IIf(optDesc.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao = '" & cboDesc.Text & "'", "")
   ElseIf chkDescPorIniciais.Value = Checked Then
      var_Criterio = Chr$(39) & cboDesc.Text & "%" & Chr(39)
      var_Criterio = IIf(optDesc.Value, IIf(var_Criterio <> "", "", " AND ") & "produtos.descricao  LIKE " & var_Criterio & "", "")
   End If
   
   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = '" & cboConsLinha.Text & "'", "")
   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = '" & txtCodBarra.Text & "'", "")
   
   If var_Criterio <> "" Then var_Criterio = " WHERE " & var_Criterio
   
   var_Indice = ""
   var_Indice = var_Indice & IIf(ckkORDDesc.Value, IIf(var_Indice <> "", ", ", "") & "produtos.descricao", "")
   var_Indice = var_Indice & IIf(ckkORDLinha.Value, IIf(var_Indice <> "", ", ", "") & "produtos.categoria", "")
   
   If var_Indice <> "" Then var_Indice = " ORDER BY " & var_Indice
   
   sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc,  produtos.ICMSaliq AS var_ICMSAliq, produtos.PISCST AS var_PIS, produtos.pisaliq AS var_PisAliq, produtos.COFINSCST AS var_COFINS, produtos.COFINSALIQ AS var_COFINSALIQ, produtos.IPICST AS var_IPI, produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST " & _
      "FROM produtos " & var_Criterio & " " & var_Indice
   
   Set r = dbData.OpenRecordset(sSQL)
   lblQuant.Caption = r.RecordCount
   
       If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        Formatar_Grid r
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
End Sub

Private Sub ckkORDDesc_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDLinha_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDQuant_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDQuantMin_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDValor_Click()
   MostrarCriterios
End Sub

Private Sub cmdAtualizar_Click()
Dim i As Integer
Dim sSQL As String

picAguarde.Visible = True
DoEvents

For i = 1 To Grid.Rows - 1
   'Atualiza a tabela de produtos
   sSQL = "UPDATE produtos SET " & _
      "cod_barra = '" & Grid.TextMatrix(i, 3) & "', " & _
      "descricao = '" & Grid.TextMatrix(i, 4) & "', " & _
      "NCM = '" & Grid.TextMatrix(i, 6) & "', " & _
      "CFOP = " & Grid.TextMatrix(i, 7) & ", " & _
      "ICMSCST = '" & Grid.TextMatrix(i, 8) & "', " & _
      "ICMSaliq = " & Replace(CDbl(Grid.TextMatrix(i, 9)), ",", ".") & ", " & _
      "pisCST = '" & Grid.TextMatrix(i, 10) & "', " & _
      "pisaliq = " & Replace(CDbl(Grid.TextMatrix(i, 11)), ",", ".") & ", " & _
      "cofinsCST = '" & Grid.TextMatrix(i, 12) & "', " & _
      "cofinsaliq = " & Replace(CDbl(Grid.TextMatrix(i, 13)), ",", ".") & ", " & _
      "ipiCST = '" & Grid.TextMatrix(i, 14) & "', " & _
      "ipialiq = " & Replace(CDbl(Grid.TextMatrix(i, 15)), ",", ".") & ", " & _
      "cest = '" & Grid.TextMatrix(i, 16) & "' " & _
      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
      'Debug.Print sSQL
   dbData.Execute sSQL
Next

picAguarde.Visible = False
End Sub

Private Sub cmdConsultaNCMean_Click()
Dim varNomeProduto As String
varNomeProduto = Grid.TextMatrix(Grid.Row, 3)
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultarNCM_Click()
Dim varNomeProduto As String
varNomeProduto = Replace(Grid.TextMatrix(Grid.Row, 4), " ", "+")
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdLocalizar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varTipoMostrar As String
Dim vUltimoValorVenda As String

If optMostrarQuant.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque > 0"
ElseIf optMostrarNegativos.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque < 0"
ElseIf optMostrarZerados.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque = 0"
ElseIf optMostrarTodos.Value = True Then
    varTipoMostrar = " "
End If

If optComPreco.Value = True Then
    vUltimoValorVenda = " and (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) > 0 "
ElseIf optSemPreco.Value = True Then
    vUltimoValorVenda = " and (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) = 0"
ElseIf optTodosPreco.Value = True Then
    vUltimoValorVenda = " and 1=1 "
End If

If optTodos.Value = True Then
   sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc,  produtos.ICMSaliq AS var_ICMSAliq, produtos.PISCST AS var_PIS, produtos.pisaliq AS var_PisAliq, produtos.COFINSCST AS var_COFINS, produtos.COFINSALIQ AS var_COFINSALIQ, produtos.IPICST AS var_IPI, produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) " & varTipoMostrar & " " & vUltimoValorVenda & " ORDER BY produtos.descricao;"

   Set r = dbData.OpenRecordset(sSQL)
   lblQuant.Caption = r.RecordCount
   
    If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        Formatar_Grid r
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
Else
   MostrarCriterios
End If

If optCodBarra.Value = True Then txtCodBarra_GotFocus
End Sub

Private Sub Form_Activate()
cmdLocalizar_Click
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
End Sub

Private Sub cboDesc_Change()
   'cboDesc_Click
End Sub

Private Sub cboDesc_Click()
   'If chkDescPorProduto.Value = Checked Then
   '   If cboDesc.Text = "" Then Exit Sub
   '   MostrarCriterios
   'ElseIf chkDescPorIniciais.Value = Checked Then
   '   MostrarCriterios
   'End If
End Sub

Private Sub cboDesc_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If chkDescPorProduto.Value = Checked Then
      cboDesc.Clear
      
      sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
      Set r = dbData.OpenRecordset(sSQL)
      
      Do While Not r.EOF
         cboDesc.AddItem ValidateNull(r("descricao"))
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      moCombo.AttachTo cboDesc
   End If
End Sub

Private Sub cboDesc_LostFocus()
   'cboDesc_Click
End Sub

Private Sub cboConsLinha_Click()
   'If cboConsLinha.Text <> "" Then MostrarCriterios
End Sub

Private Sub cboConsLinha_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboConsLinha.Clear
   
   sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsLinha.AddItem ValidateNull(r("categoria"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboConsLinha
End Sub

Private Sub cboConsLinha_LostFocus()
   'cboConsLinha_Click
End Sub

Private Sub Formatar_Grid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   LimparGrid
   picAguarde.Visible = True
   DoEvents
   
   With Grid
      .Clear
      .Cols = 17
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1300
      .ColWidth(4) = 4000
      .ColWidth(5) = 1400
      .ColWidth(6) = 850
      .ColWidth(7) = 700
      .ColWidth(8) = 700
      .ColWidth(9) = 700
      .ColWidth(10) = 650
      .ColWidth(11) = 700
      .ColWidth(12) = 800
      .ColWidth(13) = 700
      .ColWidth(14) = 650
      .ColWidth(15) = 700
      .ColWidth(16) = 900
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "NCM."
      .TextMatrix(0, 7) = "CFOP."
      .TextMatrix(0, 8) = "ICMS."
      .TextMatrix(0, 9) = "ALIQ."
      .TextMatrix(0, 10) = "PIS"
      .TextMatrix(0, 11) = "ALIQ."
      .TextMatrix(0, 12) = "COFINS"
      .TextMatrix(0, 13) = "ALIQ."
      .TextMatrix(0, 14) = "IPI"
      .TextMatrix(0, 15) = "ALIQ."
      .TextMatrix(0, 16) = "CEST"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(2) = 1
            
            '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("var_cod"))
            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("var_NCM"))
            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_CFOP"))
            .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("var_ICMS"))
            .TextMatrix(.Rows - 1, 9) = Format$(ValidateNull(rTabela("var_ICMSALIQ")), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = ValidateNull(rTabela("var_PIS"))
            .TextMatrix(.Rows - 1, 11) = Format$(ValidateNull(rTabela("var_PISALIQ")), ocMONEY)
            .TextMatrix(.Rows - 1, 12) = ValidateNull(rTabela("var_COFINS"))
            .TextMatrix(.Rows - 1, 13) = Format$(ValidateNull(rTabela("var_COFINSALIQ")), ocMONEY)
            .TextMatrix(.Rows - 1, 14) = ValidateNull(rTabela("var_IPI"))
            .TextMatrix(.Rows - 1, 15) = Format$(ValidateNull(rTabela("var_IPIALIQ")), ocMONEY)
            .TextMatrix(.Rows - 1, 16) = ValidateNull(rTabela("var_CEST"))
            
            '.TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_quant"))
            '.TextMatrix(.Rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      picAguarde.Visible = False
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Click()
Dim i As Integer

For i = 6 To 16
   If Grid.ColSel = i Then
      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
      txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)
      txtEdit.Visible = True
      txtEdit.SetFocus
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(txtEdit.Text)
      iRow = Grid.Row
      iCol = Grid.Col
   End If
Next
End Sub

Private Sub optCategoria_Click()
lblCategoria.Visible = True
cboConsLinha.Visible = True
lblDesc.Visible = False
cboDesc.Visible = False
cboDesc.Visible = False
chkDescPorProduto.Visible = False
chkDescPorIniciais.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
cboConsLinha.SetFocus
End Sub

Private Sub optCodBarra_Click()
lblCategoria.Visible = False
cboConsLinha.Visible = False
lblDesc.Visible = False
cboDesc.Visible = False
cboDesc.Visible = False
chkDescPorProduto.Visible = False
chkDescPorIniciais.Visible = False
lblCodBarra.Visible = True
txtCodBarra.Visible = True
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
txtCodBarra.SetFocus
End Sub

Private Sub optComPreco_Click()
cmdLocalizar_Click
End Sub

Private Sub optDesc_Click()
lblCategoria.Visible = False
cboConsLinha.Visible = False
lblDesc.Visible = True
cboDesc.Visible = True
chkDescPorProduto.Visible = True
chkDescPorIniciais.Visible = True
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
cboDesc.SetFocus
End Sub

Private Sub optMostrarNegativos_Click()
cmdLocalizar_Click
End Sub

Private Sub optMostrarQuant_Click()
cmdLocalizar_Click
End Sub

Private Sub optMostrarTodos_Click()
cmdLocalizar_Click
End Sub

Private Sub optMostrarZerados_Click()
cmdLocalizar_Click
End Sub

Private Sub optSemPreco_Click()
cmdLocalizar_Click
End Sub

Private Sub optTodos_Click()
lblCategoria.Visible = False
cboConsLinha.Visible = False
lblDesc.Visible = False
cboDesc.Visible = False
cboDesc.Visible = False
chkDescPorProduto.Visible = False
chkDescPorIniciais.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = False
Frame6.Enabled = True
optComPreco.Value = True
cmdLocalizar_Click
End Sub

Private Sub optTodosPreco_Click()
cmdLocalizar_Click
End Sub

Private Sub txtCodBarra_Change()
   If Len(txtCodBarra.Text) = 13 Then cmdLocalizar_Click
End Sub

Private Sub txtCodBarra_GotFocus()
   SelectControl txtCodBarra
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   'Exit Sub
   If KeyCode = 38 Then
      If Grid.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
      Grid.Row = iRow - 1
      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      Grid_Click
   
   ElseIf KeyCode = 40 Then
      If Grid.Rows = Grid.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      Grid.Row = iRow + 1
      Grid_Click
   End If
End Sub

Private Sub txtEdit_LostFocus()

If iCol = 6 Then
    txtEdit.Text = Replace(txtEdit.Text, ".", "")
    txtEdit.Text = Trim(txtEdit.Text)
ElseIf iCol = 9 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 11 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 13 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 15 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 16 Then
    txtEdit.Text = Replace(txtEdit.Text, ".", "")
    txtEdit.Text = Trim(txtEdit.Text)
End If
Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)

txtEdit.Visible = False
End Sub



