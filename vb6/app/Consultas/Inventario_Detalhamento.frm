VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Inventario_Detalhamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE INVENTÁRIO"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15180
   Icon            =   "Inventario_Detalhamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Totais"
      Height          =   675
      Left            =   12180
      TabIndex        =   20
      Top             =   8400
      Width           =   2895
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Total Fiscal:"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblValorTotalFiscal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1260
         TabIndex        =   21
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame frmCriterio 
      Caption         =   "Tipo de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   16
      Top             =   8400
      Width           =   3045
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   540
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame frmExibir 
      Caption         =   "Exibir"
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
      Height          =   675
      Left            =   10500
      TabIndex        =   13
      Top             =   8400
      Width           =   1335
      Begin VB.CheckBox chkZerado 
         Caption         =   "Zerado"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkNegativo 
         Caption         =   "Negativo"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   1035
      End
   End
   Begin VB.Frame frmIncluir 
      Caption         =   "Incluir"
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
      Height          =   675
      Left            =   9000
      TabIndex        =   7
      Top             =   8400
      Width           =   1455
      Begin VB.CheckBox chkConsumo 
         Caption         =   "Consumo"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   1035
      End
      Begin VB.CheckBox chkImobilizado 
         Caption         =   "Imobilizados"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3120
      TabIndex        =   5
      Top             =   8400
      Width           =   5865
      Begin VB.ComboBox cboDesc 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   3555
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Consultar"
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
         MICON           =   "Inventario_Detalhamento.frx":1D82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCategoria 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Inventario_Detalhamento.frx":1D9E
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
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   60
      Width           =   15015
      Begin VB.Label lblExercicio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "EXERCICIO XXXX"
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
         Left            =   13320
         TabIndex        =   19
         Top             =   360
         Width           =   1545
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   480
         Picture         =   "Inventario_Detalhamento.frx":2DD6
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTÁRIO"
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
         Width           =   1860
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   9135
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22437
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "20:59"
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
      Height          =   7275
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12832
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
End
Attribute VB_Name = "Inventario_Detalhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim r As ADODB.Recordset
Dim printSQL As String
Private moCombo As cComboHelper

Public Sub LerInventario(ByVal vIDInvent As Long)
'sSQL = "SELECT TbInventariosItens.Seq as vSeq, TbInventariosItens.IDProduto as vCodProd, TbInventariosItens.NomeProduto as vDesc, produtos.EAN as vEAN , produtos.NCM as vNCM,  TbInventariosItens.MetaCalculado as vMeta, TbInventariosItens.VlrUnitInvent as vVlrUnit, TbInventariosItens.TotalInvent as vTotal " & _
       "FROM TbInventariosItens INNER JOIN produtos ON TbInventariosItens.IDProduto = produtos.CODIGO " & _
       "WHERE IdInventario = " & vIDInvent & " " & _
       "ORDER BY TbInventariosItens.Seq"
sSQL = "SELECT IdInventario, Seq as vSeq, IDProduto as vCodProd, NomeProduto as vDesc, EAN as vEAN, NCM as vNCM, SaldoCalculado, VlrUnitInvent as vVlrUnit, TotalInvent as vTotal, MetaCalculado  as vMeta, SaldoLancado, TotalFisico, AnoInventario " & _
       "FROM InventarioGerado " & _
       "WHERE IdInventario = " & vIDInvent & " " & _
       "ORDER BY Seq"
       'IIf(Not Vazio(Filtro), "WHERE " & Filtro, "") & " " & _
       'IIf(Not SemOrderBy, "ORDER BY Seq", "")
       'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

lblExercicio.Caption = "EXERCÍCIO: " & ValidateNull(r("AnoInventario"))

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Private Sub cboDesc_GotFocus()
cboDesc.Clear

sSQL = "SELECT DISTINCT NomeProduto FROM InventarioGerado ORDER BY NomeProduto;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboDesc.AddItem ValidateNull(r("NomeProduto"))
   r.MoveNext
Loop

SelectControl cboDesc
moCombo.AttachTo cboDesc
End Sub


Private Sub cboTipo_GotFocus()
cboTipo.Clear
cboTipo.AddItem "FISICO"
cboTipo.AddItem "FISCAL"
moCombo.AttachTo cboTipo
End Sub





Private Sub cmdLocalizar_Click()
If cboDesc.Text = "" Then
    sSQL = "SELECT IdInventario, Seq as vSeq, IDProduto as vCodProd, NomeProduto as vDesc, EAN as vEAN, NCM as vNCM, SaldoCalculado, VlrUnitInvent as vVlrUnit, TotalInvent as vTotal, MetaCalculado  as vMeta, SaldoLancado, TotalFisico " & _
           "FROM InventarioGerado " & _
           "WHERE 1=0  " & _
           "ORDER BY Seq"
Else
    sSQL = "SELECT IdInventario, Seq as vSeq, IDProduto as vCodProd, NomeProduto as vDesc, EAN as vEAN, NCM as vNCM, SaldoCalculado, VlrUnitInvent as vVlrUnit, TotalInvent as vTotal, MetaCalculado  as vMeta, SaldoLancado, TotalFisico " & _
           "FROM InventarioGerado " & _
           "WHERE (NomeProduto = '" & cboDesc.Text & "')  " & _
           "ORDER BY Seq"
End If
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

End Sub



Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   
   .ColWidth(1) = 750
   .ColWidth(2) = 850
   .ColWidth(3) = 1600
   .ColWidth(4) = 5000
   .ColWidth(5) = 1200
   .ColWidth(6) = 1500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   
   .TextMatrix(0, 1) = "SEQ"
   .TextMatrix(0, 2) = "CÓDIGO"
   .TextMatrix(0, 3) = "EAN"
   .TextMatrix(0, 4) = "DESCRIÇÃO"
   .TextMatrix(0, 5) = "NCM"
   .TextMatrix(0, 6) = "META."
   .TextMatrix(0, 7) = "UNIT"
   .TextMatrix(0, 8) = "TOTAL"

   .Redraw = False
   i = 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("vSeq")
         .TextMatrix(.rows - 1, 2) = Format(ValidateNull(rTabela("vcodprod")), "00000")
         .TextMatrix(.rows - 1, 3) = rTabela("vEAN")
         .TextMatrix(.rows - 1, 4) = rTabela("vdesc")
         .TextMatrix(.rows - 1, 5) = rTabela("vNCM")
         .TextMatrix(.rows - 1, 6) = Format(rTabela("vmeta"), ocMONEY)
         .TextMatrix(.rows - 1, 7) = Format(rTabela("vVlrUnit"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = Format(rTabela("vtotal"), ocMONEY)
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If

   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   
   'MUDAR COR DE FONTE DA COLUNA
   'For i = 1 To .Rows - 1
   '   .Row = i
   '   .Col = 9
   '   .CellForeColor = &HC0&
   '   .CellFontBold = True
   'Next
   
   .rows = .rows - 1
   .Redraw = True
End With

lblValorTotalFiscal.Caption = Format(SomaGrid(Grid, 8), ocMONEY)

'lblTotalEntrada.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
'lblTotalSaida.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
'lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End Sub
Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Function Grid1Filtro() As String
'    Dim FI As String, ComandoSQL As String
    
    'Filtro Fixo
'    cf FI, "IdInventario = " & xIdInvent
    
    'Tipo
'    If Not Vazio(cboTipo.Text) Then
'      If cboTipo.Value = "FÍSICO" Then
'        If Not ckMostrarErros Then
'          cf FI, "SaldoLancado > 0"
'        Else
'          cf FI, "SaldoLancado <= 0"
'        End If
'      Else
'        If Not ckMostrarErros Then
'          cf FI, "MetaCalculado > 0"
'        Else
'          cf FI, "MetaCalculado <= 0"
'        End If
'      End If
'    Else
'      cf FI, "1=2"
'    End If
    
'    If Not Imobilizado And Not UsoConsumo Then
'       cf FI, "Imobilizado = 0"
'       cf FI, "UsoConsumo = 0"
'    ElseIf Not Imobilizado And UsoConsumo Then
'       cf FI, "Imobilizado = 0"
'    ElseIf Imobilizado And Not UsoConsumo Then
'       cf FI, "UsoConsumo = 0"
'    End If

    'NomeProduto
'    If Not Vazio(cboDesc.Text) Then cf FI, "NomeProduto LIKE '" & cboDesc.Text & "%'"

    'Retorna
'    Grid1Filtro = FI
End Function

Private Function Grid1SQL(Optional SemOrderBy As Boolean = False, Optional InsertInto As Boolean = False) As String
Dim Filtro As String, r As String
'Pega o Filtro
Filtro = Grid1Filtro()
'SQL
If InStr(Filtro, "1=2") > 0 Then
   r = "SELECT 0 As IdInventario, 0 As Seq, 0 As IDProduto, 'INFORME O CRITERIO PARA PESQUISAR' As NomeProduto, '' As EAN, '' As NCM, 0 As Saldo, 0 As VlrUnitInvent, 0 As TotalInvent, 0 As Meta, 0 As SaldoLancado, 0 As TotalFisico " & _
       "FROM TbEmpresa"
Else
   r = "SELECT IdInventario, Seq, IDProduto, NomeProduto, EAN, NCM, SaldoCalculado, VlrUnitInvent, TotalInvent, MetaCalculado, SaldoLancado, TotalFisico " & _
       "FROM InventarioGerado " & _
       IIf(Not Vazio(Filtro), "WHERE " & Filtro, "") & " " & _
       IIf(Not SemOrderBy, "ORDER BY Seq", "")
End If
'Faz retorno
Grid1SQL = r
End Function

Private Sub Form_Load()
Set moCombo = New cComboHelper
End Sub


