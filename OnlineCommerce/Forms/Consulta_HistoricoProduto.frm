VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Consulta_HistoricoProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HISTÓRICO DO PRODUTO"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "Consulta_HistoricoProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   10485
      TabIndex        =   34
      Top             =   60
      Width           =   10515
      Begin VB.Label lblInfoProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione um produto para ver o histórico."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Consulta_HistoricoProduto.frx":23D2
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LINHA DO TEMPO DO PRODUTO"
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
         Left            =   1620
         TabIndex        =   35
         Top             =   120
         Width           =   4860
      End
   End
   Begin VB.Frame Frame_Criterios 
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
      Height          =   2115
      Left            =   60
      TabIndex        =   19
      Top             =   1020
      Width           =   2295
      Begin VB.ComboBox cboCriterioSec 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1995
      End
      Begin VB.ComboBox cboCriterioPrinc 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Consultar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   880
         Width           =   690
      End
   End
   Begin VB.Frame Frame_Filtros 
      Caption         =   "Filtros"
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
      Height          =   2115
      Left            =   2400
      TabIndex        =   22
      Top             =   1020
      Width           =   8175
      Begin VB.TextBox txtCodBarra 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox cboDescricao 
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.ComboBox cboForma 
         Height          =   315
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   480
         Width           =   1700
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   60
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   60
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFim 
         Height          =   315
         Left            =   1380
         TabIndex        =   27
         Top             =   1080
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario1 
         Height          =   315
         Left            =   1020
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "Calendario"
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Consulta_HistoricoProduto.frx":8C18
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario2 
         Height          =   315
         Left            =   2340
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "Calendario"
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Consulta_HistoricoProduto.frx":8C34
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdBuscar 
         Height          =   495
         Left            =   60
         TabIndex        =   12
         Top             =   1560
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Buscar"
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Consulta_HistoricoProduto.frx":8C50
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Barra:"
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblForma 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma:"
         Height          =   195
         Left            =   3060
         TabIndex        =   31
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano:"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data inicial:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblFim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data final:"
         Height          =   195
         Left            =   1380
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4200
      Left            =   60
      TabIndex        =   0
      Top             =   3180
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7408
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirNota 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   7440
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "MOSTRAR DETALHE"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Consulta_HistoricoProduto.frx":8C6C
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
      Height          =   315
      Left            =   2460
      TabIndex        =   8
      Top             =   7440
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "IMPRIMIR"
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
      MICON           =   "Consulta_HistoricoProduto.frx":8C88
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
      TabIndex        =   36
      Top             =   8190
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14446
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "19:21"
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
   Begin VB.Label lblLegendaEntrada 
      AutoSize        =   -1  'True
      Caption         =   "Preto: Entrada / Adição"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4500
      TabIndex        =   15
      Top             =   7600
      Width           =   1680
   End
   Begin VB.Label lblLegendaSaida 
      AutoSize        =   -1  'True
      Caption         =   "Vermelho: Saída / Remoção"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4500
      TabIndex        =   16
      Top             =   7400
      Width           =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remoção:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7500
      TabIndex        =   7
      Top             =   7800
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adição:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7740
      TabIndex        =   6
      Top             =   7440
      Width           =   720
   End
   Begin VB.Label lblQuantRemocao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8595
      TabIndex        =   5
      Top             =   7800
      Width           =   630
   End
   Begin VB.Label lblQuantAdicao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8610
      TabIndex        =   4
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label lblTotalRemocao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9270
      TabIndex        =   3
      Top             =   7800
      Width           =   1230
   End
   Begin VB.Label lblTotalAdicao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9270
      TabIndex        =   2
      Top             =   7440
      Width           =   1230
   End
End
Attribute VB_Name = "Consulta_HistoricoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Private moCombo As cComboHelper
'True so durante o Form_Load - impede que o ListIndex programatico dos
'combos dispare o Click (que da SetFocus) antes da janela existir (erro 5)
Private bCarregando As Boolean

'Guarda o filtro de periodo em uso na ultima busca - pra replicar o mesmo
'filtro quando cmdExibirNota abrir outra tela (menos venda, que abre a
'consulta da venda especifica, sem precisar de filtro de periodo)
Private vFiltroCriterioPrinc As String
Private vFiltroMes As Integer
Private vFiltroAno As String
Private vFiltroDataIni As String
Private vFiltroDataFim As String

'Um evento do historico, ja com o saldo acumulado calculado (ver MontarEventos)
Private Type TEvento
   DATA As Date
   hora As String
   Forma As String
   Tipo As String
   QUANT As Double
   Usuario As String
   CodEntrada As Long
   CodProduto As Long
   Saldo As Double
End Type

Private eventos() As TEvento
Private nEventos As Integer

Public Sub loadHistorico(ByVal CodProduto As Long, Optional ByVal CriterioPrinc As String = "TODOS", Optional ByVal MES As Integer = 0, Optional ByVal ANO As String = "", Optional ByVal DataIni As String = "", Optional ByVal DataFim As String = "", Optional ByVal FormaFiltro As String = "")
Dim sSQL As String
Dim r As ADODB.Recordset

vFiltroCriterioPrinc = CriterioPrinc
vFiltroMes = MES
vFiltroAno = ANO
vFiltroDataIni = DataIni
vFiltroDataFim = DataFim

sSQL = "SELECT descricao, cod_barra, quant_estoque FROM produtos WHERE (codigo = " & CodProduto & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then
   lblInfoProduto.Caption = "PRODUTO: " & ValidateNull(r("descricao"))
Else
   lblInfoProduto.Caption = "Produto não encontrado."
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Exit Sub
End If

If r.State <> 0 Then r.Close
Set r = Nothing

'Busca SEMPRE o historico INTEIRO do produto (sem filtro de periodo/forma na
'consulta SQL) - o filtro e aplicado depois, evento a evento, em
'MontarEventos. Isso garante que o SALDO acumulado de cada linha mostrada
'seja o saldo real (considerando tambem os eventos escondidos pelo filtro),
'e nao um saldo que comeca zerado no meio da historia
sSQL = "SELECT Produtos_Quant.Data AS vData, " & _
      "CASE WHEN CONVERT(VARCHAR(8), Produtos_Quant.Hora, 108) = '00:00:00' THEN '' ELSE CONVERT(VARCHAR(8), Produtos_Quant.Hora, 108) END AS vHora, " & _
      "Produtos_Quant.FORMA AS vForma, Produtos_Quant.TIPO AS vTipo, ABS(Produtos_Quant.QUANT) AS vQuant, " & _
      "ISNULL(funcionario.NOME, '') AS vUsuario, ISNULL(Produtos_Quant.COD_ENTRADA, 0) AS vCodEntrada " & _
      "FROM Produtos_Quant LEFT JOIN funcionario ON funcionario.CODIGO = Produtos_Quant.cod_usuario " & _
      "WHERE (Produtos_Quant.COD_PRODUTO = " & CodProduto & ") " & _
      "UNION ALL " & _
      "SELECT pedidos.DATA_COMPRA, " & _
      "CASE WHEN CONVERT(VARCHAR(8), pedidos.HORA_COMPRA, 108) = '00:00:00' THEN '' ELSE CONVERT(VARCHAR(8), pedidos.HORA_COMPRA, 108) END, " & _
      "'VENDA', 'REMOÇÃO', pedidos_itens.quantidade, " & _
      "ISNULL(funcionario.NOME, ''), pedidos.cod_pedido " & _
      "FROM pedidos_itens INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      "LEFT JOIN funcionario ON funcionario.CODIGO = pedidos.cod_funcionario " & _
      "WHERE (pedidos_itens.cod_produto = " & CodProduto & ") AND (pedidos.CANCELADO = 0) AND (pedidos.tipo_pedido <> 'ORÇAMENTO') " & _
      "AND (pedidos.tipo_pagamento IN ('À Vista', 'À prazo')) AND (ISNULL(pedidos_itens.tipo_venda,'') <> 'OFICINA') " & _
      "ORDER BY vData, vHora"

Set r = dbData.OpenRecordset(sSQL)

MontarEventos r, CodProduto, CriterioPrinc, MES, ANO, DataIni, DataFim, FormaFiltro

If r.State <> 0 Then r.Close
Set r = Nothing

MontarGrid
SomaTotais
End Sub

'Percorre o recordset INTEIRO (ja ordenado por data/hora) acumulando o saldo
'evento a evento - e so DEPOIS de acumular decide se aquele evento entra no
'array exibido (conforme FormaFiltro/CriterioPrinc). Assim o saldo de uma
'linha filtrada continua sendo o saldo real naquele ponto da historia, nao
'um saldo que reinicia do zero no primeiro evento do filtro
Private Sub MontarEventos(rTabela As ADODB.Recordset, ByVal CodProduto As Long, ByVal CriterioPrinc As String, ByVal MES As Integer, ByVal ANO As String, ByVal DataIni As String, ByVal DataFim As String, ByVal FormaFiltro As String)
Dim saldoAcumulado As Double
Dim n As Integer
Dim incluir As Boolean
Dim dData As Date, sForma As String, sTipo As String, dQuant As Double

ReDim eventos(200)
n = 0
saldoAcumulado = 0

Do While Not rTabela.EOF
   dData = rTabela("vData")
   sForma = rTabela("vForma")
   sTipo = rTabela("vTipo")
   dQuant = rTabela("vQuant")

   If sTipo = "ADIÇÃO" Then
      saldoAcumulado = saldoAcumulado + dQuant
   Else
      saldoAcumulado = saldoAcumulado - dQuant
   End If

   incluir = True
   If FormaFiltro <> "" Then
      If sForma <> FormaFiltro Then incluir = False
   End If
   If incluir And CriterioPrinc = "MENSAL" Then
      If Month(dData) <> MES Or CStr(Year(dData)) <> ANO Then incluir = False
   ElseIf incluir And CriterioPrinc = "PERÍODO" Then
      If dData < CDate(DataIni) Or dData > CDate(DataFim) Then incluir = False
   End If

   If incluir Then
      If n > UBound(eventos) Then ReDim Preserve eventos(UBound(eventos) + 200)

      eventos(n).DATA = dData
      eventos(n).hora = ValidateNull(rTabela("vHora"))
      eventos(n).Forma = sForma
      eventos(n).Tipo = sTipo
      eventos(n).QUANT = dQuant
      eventos(n).Usuario = ValidateNull(rTabela("vUsuario"))
      eventos(n).CodEntrada = Val(ValidateNull(rTabela("vCodEntrada")))
      eventos(n).CodProduto = CodProduto
      eventos(n).Saldo = saldoAcumulado

      n = n + 1
   End If

   rTabela.MoveNext
Loop

nEventos = n
End Sub

'Totais do rodape (Adicao/Remocao) - calculado direto dos eventos, nao do
'texto do grid (que tem '+'/'-' e '(Nx)' quando agrupado)
Private Sub SomaTotais()
Dim i As Integer
Dim somaAdicao As Double, somaRemocao As Double
Dim qtdAdicao As Integer, qtdRemocao As Integer

For i = 0 To nEventos - 1
   If eventos(i).Tipo = "ADIÇÃO" Then
      somaAdicao = somaAdicao + eventos(i).QUANT
      qtdAdicao = qtdAdicao + 1
   Else
      somaRemocao = somaRemocao + eventos(i).QUANT
      qtdRemocao = qtdRemocao + 1
   End If
Next

lblTotalAdicao.Caption = somaAdicao
lblQuantAdicao.Caption = Format(qtdAdicao, "000")
lblTotalRemocao.Caption = somaRemocao
lblQuantRemocao.Caption = Format(qtdRemocao, "000")
End Sub

'Monta o grid - sem agrupamento, cada evento na sua propria linha (ver
'MontarEventos e AdicionarLinha)
Private Sub MontarGrid()
Dim i As Integer

With Grid
   .Clear
   .Cols = 8
   .rows = 2

   .ColWidth(0) = 900
   .ColWidth(1) = 800
   .ColWidth(2) = 1360
   .ColWidth(3) = 1000
   .ColWidth(4) = 1300
   .ColWidth(5) = 1000
   .ColWidth(6) = 1300
   .ColWidth(7) = 0

   .TextMatrix(0, 0) = "DATA"
   .TextMatrix(0, 1) = "HORA"
   .TextMatrix(0, 2) = "USUÁRIO"
   .TextMatrix(0, 3) = "TIPO"
   .TextMatrix(0, 4) = "FORMA"
   .TextMatrix(0, 5) = "QUANT"
   .TextMatrix(0, 6) = "SALDO"

   .Redraw = False

   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
      .CellAlignment = flexAlignCenterCenter
   Next

   .ColAlignment(0) = 4
   .ColAlignment(1) = 4
   .ColAlignment(5) = 7
   .ColAlignment(6) = 7

   If nEventos = 0 Then
      .rows = .rows - 1
      .Redraw = True
      Exit Sub
   End If

   'Sem agrupamento - cada evento fica na sua propria linha, com a data e
   'hora dele, pra nao embaralhar a ordem cronologica real
   For i = 0 To nEventos - 1
      AdicionarLinha i
   Next

   .rows = .rows - 1
   .Row = 0
   .Col = 0
   .Redraw = True
End With
End Sub

Private Sub AdicionarLinha(ByVal idx As Integer)
With Grid
   .TextMatrix(.rows - 1, 0) = Format(eventos(idx).DATA, "dd/mm/yy")
   .TextMatrix(.rows - 1, 1) = eventos(idx).hora
   .TextMatrix(.rows - 1, 2) = eventos(idx).Usuario
   .TextMatrix(.rows - 1, 3) = eventos(idx).Tipo
   .TextMatrix(.rows - 1, 4) = eventos(idx).Forma
   .TextMatrix(.rows - 1, 5) = FormatarQuant(eventos(idx).QUANT)
   .TextMatrix(.rows - 1, 6) = Format(eventos(idx).Saldo, "0")
   .TextMatrix(.rows - 1, 7) = eventos(idx).CodEntrada

   ColorirLinha .rows - 1, eventos(idx).Tipo

   .rows = .rows + 1
End With
End Sub

Private Function FormatarQuant(ByVal QUANT As Double) As String
FormatarQuant = Format(QUANT, "0")
End Function

'Preto pra ADIÇÃO (entrada), vermelho escuro pra REMOÇÃO (saida)
Private Sub ColorirLinha(ByVal linha As Integer, ByVal Tipo As String)
Dim j As Integer
Dim corTexto As Long

If Tipo = "ADIÇÃO" Then
   corTexto = vbBlack
Else
   corTexto = RGB(128, 0, 0)
End If

With Grid
   For j = 0 To .Cols - 1
      .Row = linha
      .Col = j
      .CellForeColor = corTexto
   Next
   .Row = linha
   .Col = 6
   .CellBackColor = &HC0FFFF
End With
End Sub

Private Sub cmdBuscar_Click()
Dim lCod As Long
Dim sFormaFiltro As String

If cboCriterioSec.Text = "CÓD. BARRA" Then
   If Trim(txtCodBarra.Text) = "" Then
      ShowMsg "Digite o código de barra.", vbInformation
      Exit Sub
   End If
   lCod = Val(SQLExecutaRetorno("SELECT codigo FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "')", "codigo", "0"))
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" Then
   If Trim(txtCodBarra.Text) = "" Or Not IsNumeric(txtCodBarra.Text) Then
      ShowMsg "Digite o código do produto.", vbInformation
      Exit Sub
   End If
   lCod = Val(txtCodBarra.Text)
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   If cboDescricao.ListIndex = -1 Then
      ShowMsg "Escolha a descrição do produto.", vbInformation
      Exit Sub
   End If
   lCod = cboDescricao.ItemData(cboDescricao.ListIndex)
Else
   Exit Sub
End If

If lCod = 0 Then
   ShowMsg "Produto não encontrado.", vbInformation
   Exit Sub
End If

sFormaFiltro = ""
If cboForma.Text <> "" And cboForma.Text <> "TODOS" Then sFormaFiltro = cboForma.Text

If cboCriterioPrinc.Text = "MENSAL" Then
   If cboAno.Text = "" Or cboMes.Text = "" Then
      ShowMsg "Escolha o mês e o ano.", vbInformation
      Exit Sub
   End If
   loadHistorico lCod, "MENSAL", cboMes.ListIndex + 1, cboAno.Text, , , sFormaFiltro
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
   If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then
      ShowMsg "Digite as duas datas do período.", vbInformation
      Exit Sub
   End If
   loadHistorico lCod, "PERÍODO", , , mskInicio.Text, mskFim.Text, sFormaFiltro
Else
   loadHistorico lCod, "TODOS", , , , , sFormaFiltro
End If
End Sub

Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdBuscar_Click
End Sub

Private Sub cboDescricao_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vTexto As String

SelectControl cboDescricao

vTexto = cboDescricao.Text
cboDescricao.Clear

sSQL = "SELECT codigo, descricao FROM produtos ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboDescricao.AddItem ValidateNull(r("descricao"))
   cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboDescricao.Text = vTexto
moCombo.AttachTo cboDescricao
End Sub

Private Sub cboDescricao_LostFocus()
If cboDescricao.ListIndex <> -1 Then txtCodBarra.Text = ""
End Sub

'Linha selecionada: FORMA=ENTRADA abre a nota fiscal exata; CADASTRO/AJUSTE/
'SAÍDA/OFICINA abre o histórico de ajuste desse produto; VENDA abre a
'consulta de vendas desse produto. Mesmo padrão de roteamento já usado
'nas outras telas que mexemos esse mes.
Private Sub cmdExibirNota_Click()
If Grid.Row = 0 Then Exit Sub
If nEventos = 0 Then Exit Sub

Dim sForma As String
sForma = Grid.TextMatrix(Grid.Row, 4)

Select Case sForma
   Case "ENTRADA"
      'Veio de nota fiscal - abre direto a nota exata (coluna oculta 7 = COD_ENTRADA)
      If IsNumeric(Grid.TextMatrix(Grid.Row, 7)) Then
         If Val(Grid.TextMatrix(Grid.Row, 7)) <> 0 Then
            Produtos_Entrada.AbrirNotaDireto Val(Grid.TextMatrix(Grid.Row, 7))
            Produtos_Entrada.Show
         End If
      End If

   Case "CADASTRO", "AJUSTE", "CONTAGEM"
      'Replica o mesmo filtro de periodo que estava ativo aqui (TODOS/MENSAL/
      'PERIODO) - FiltrarPorProdutoEPeriodo nao marca volta pra nenhuma tela
      'Agrupada, diferente do FiltrarPorProdutoEForma usado em outro lugar
      Ajuste_Consulta_PorProdutos.FiltrarPorProdutoEPeriodo eventos(0).CodProduto, vFiltroCriterioPrinc, vFiltroMes, vFiltroAno, vFiltroDataIni, vFiltroDataFim
      Ajuste_Consulta_PorProdutos.Show

   Case "SAÍDA"
      'Não usa FiltrarPorProdutoEForma aqui de propósito - aquele método
      'marca a tela pra voltar pro Saida_Consulta_PorProdutosAgrupadas quando
      'fechar, o que não faz sentido vindo desse histórico. FiltrarPorProdutoEPeriodo
      'e a versao sem essa marca, e ja replica o filtro de periodo
      Saida_Consulta_PorProdutos.FiltrarPorProdutoEPeriodo eventos(0).CodProduto, "SAÍDA", vFiltroCriterioPrinc, vFiltroMes, vFiltroAno, vFiltroDataIni, vFiltroDataFim
      Saida_Consulta_PorProdutos.Show

   Case "VENDA"
      'coluna oculta 7 guarda o cod_pedido de verdade pra venda (ver
      'loadHistorico - reaproveita a mesma coluna que guarda COD_ENTRADA)
      If IsNumeric(Grid.TextMatrix(Grid.Row, 7)) Then
         If Val(Grid.TextMatrix(Grid.Row, 7)) <> 0 Then
            Parcelas_Consulta_Produtos.loadPedidos Val(Grid.TextMatrix(Grid.Row, 7)), "VENDA"
            Parcelas_Consulta_Produtos.Show 1
         End If
      End If

   Case Else
      'OFICINA (consumo de peça na OS) ou qualquer forma nova - ainda não
      'tem uma tela de detalhe dedicada
      ShowMsg "Esse tipo de movimento (" & sForma & ") ainda não tem uma tela de detalhe específica.", vbInformation
End Select
End Sub

Private Sub cmdImprimir_Click()
If nEventos = 0 Then
   ShowMsg "Nada para imprimir - faça uma busca primeiro.", vbInformation
   Exit Sub
End If

Dim sDescricaoProduto As String
sDescricaoProduto = SQLExecutaRetorno("SELECT descricao FROM produtos WHERE (codigo = " & eventos(0).CodProduto & ")", "descricao", "")

'Essa tela nao guarda um printSQL como as outras 3 (o filtro de periodo/
'forma e aplicado em VB, nao em SQL - ver MontarEventos) - entao o jeito de
'garantir que o relatorio mostre exatamente o que esta na tela e montar um
'recordset em memoria a partir do proprio array eventos() ja filtrado
Dim rPrint As ADODB.Recordset
Set rPrint = New ADODB.Recordset
rPrint.CursorLocation = adUseClient
rPrint.Fields.Append "DATA", adDate
rPrint.Fields.Append "HORA", adDate
rPrint.Fields.Append "varNomeUsuario", adVarChar, 100
rPrint.Fields.Append "DESCRICAO", adVarChar, 200
rPrint.Fields.Append "FORMA", adVarChar, 20
rPrint.Fields.Append "TIPO", adVarChar, 20
rPrint.Fields.Append "QUANT", adDouble
rPrint.Open

Dim i As Integer
For i = 0 To nEventos - 1
   rPrint.AddNew
   rPrint("DATA") = eventos(i).DATA
   If eventos(i).hora <> "" Then rPrint("HORA") = CDate(eventos(i).hora)
   rPrint("varNomeUsuario") = eventos(i).Usuario
   rPrint("DESCRICAO") = sDescricaoProduto
   rPrint("FORMA") = eventos(i).Forma
   rPrint("TIPO") = eventos(i).Tipo
   rPrint("QUANT") = eventos(i).QUANT
   rPrint.Update
Next
rPrint.MoveFirst

'So tem 3 quadradinhos de total no relatorio compartilhado (rotulos fixos
'CADASTRO/ENTRADA/AJUSTE), mas essa tela tem 6 tipos de movimento - por
'pedido do usuario, usa ADIÇÃO no 1º quadro e REMOÇÃO no 2º (mesmos totais que
'ja aparecem no rodape da tela - lblTotalAdicao/lblTotalRemocao), e deixa o
'3º sem uso. O TOTAL geral (abaixo) continua somando tudo, incluindo
'SAÍDA/VENDA/OFICINA que não tem quadradinho proprio
Dim somaTotalGeral As Double
For i = 0 To nEventos - 1
   somaTotalGeral = somaTotalGeral + eventos(i).QUANT
Next

Me.Hide

Set REL_Entrada_PorProduto.Relatorio.Recordset = rPrint

REL_Entrada_PorProduto.lblTitulo.Caption = "RELATÓRIO DE HISTÓRICO DO PRODUTO"

'Os rotulos dos 2 quadradinhos usados viram ADICAO/REMOCAO (em vez de
'CADASTRO/ENTRADA, que nao fazem sentido aqui) - Unload no final restaura
'tudo isso pro proximo print, seja dessa tela ou de outra
REL_Entrada_PorProduto.Label6.Caption = "ADIÇÃO:"
REL_Entrada_PorProduto.Label7.Caption = "REMOÇÃO:"

REL_Entrada_PorProduto.dfQuantTotal.Caption = somaTotalGeral & "  "
REL_Entrada_PorProduto.dfQuantCadastroSoma.Caption = lblTotalAdicao.Caption & "  "
REL_Entrada_PorProduto.dfQuantEntradaSoma.Caption = lblTotalRemocao.Caption & "  "
REL_Entrada_PorProduto.dfQuantAjusteSoma.Caption = "0  "

REL_Entrada_PorProduto.dfQuantCadastro.Caption = lblQuantAdicao.Caption & "  "
REL_Entrada_PorProduto.dfQuantEntrada.Caption = lblQuantRemocao.Caption & "  "
REL_Entrada_PorProduto.dfQuantAjuste.Caption = Format(0, "000") & "  "

REL_Entrada_PorProduto.rfCons1.Caption = "PRODUTO = " & sDescricaoProduto

If vFiltroCriterioPrinc = "MENSAL" Then
   REL_Entrada_PorProduto.rfCons2.Caption = "Mês/Ano = " & vFiltroMes & "/" & vFiltroAno
ElseIf vFiltroCriterioPrinc = "PERÍODO" Then
   REL_Entrada_PorProduto.rfCons2.Caption = "Intervalo de " & vFiltroDataIni & " à " & vFiltroDataFim
Else
   REL_Entrada_PorProduto.rfCons2.Caption = "TODOS"
End If

'O 3 quadrado (AJUSTE) fica sem uso nessa tela - Visible precisa ser
'setado por ULTIMO, depois de mexer no Caption dos campos, porque
'setar Caption nesse controle (ReportX.ReportField) reativa o Visible.
'Tambem zera Width/Height - Visible sozinho pode nao bastar pra sumir
'com a borda (Borda=15) desenhada pela secao do relatorio. O Unload
'no final restaura tudo (tamanho/visibilidade/rotulo) pro proximo print
REL_Entrada_PorProduto.Label11.Visible = False
REL_Entrada_PorProduto.dfQuantAjuste.Visible = False
REL_Entrada_PorProduto.dfQuantAjuste.Width = 0
REL_Entrada_PorProduto.dfQuantAjuste.Height = 0
REL_Entrada_PorProduto.dfQuantAjusteSoma.Visible = False
REL_Entrada_PorProduto.dfQuantAjusteSoma.Width = 0
REL_Entrada_PorProduto.dfQuantAjusteSoma.Height = 0

REL_Entrada_PorProduto.Relatorio.Ativar
Unload REL_Entrada_PorProduto

Me.Show
End Sub

Private Sub PreencherConsulta()
cboCriterioSec.AddItem "DESCRIÇÃO"
cboCriterioSec.AddItem "CÓD. BARRA"
cboCriterioSec.AddItem "CÓD. PRODUTO"
End Sub

Private Sub PreencherCriterios()
cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "PERÍODO"
End Sub

Private Sub AtualizarCriterioSec()
If cboCriterioSec.Text = "" Then cboCriterioSec.Text = "CÓD. BARRA"

If cboCriterioSec.Text = "CÓD. BARRA" Then
   lblDescricao.Caption = "Cód. Barra:"
   txtCodBarra.Visible = True
   txtCodBarra.Text = ""
   cboDescricao.Visible = False
   cboForma.Left = 3060
   lblForma.Left = 3060
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" Then
   lblDescricao.Caption = "Cód. Produto:"
   txtCodBarra.Visible = True
   txtCodBarra.Text = ""
   cboDescricao.Visible = False
   cboForma.Left = 3060
   lblForma.Left = 3060
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   lblDescricao.Caption = "Descrição:"
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   cboForma.Left = 6120
   lblForma.Left = 6120
End If
End Sub

Private Sub cboCriterioSec_Click()
'Não espera perder o foco - exibe os objetos e dá foco na hora que escolhe.
'SetFocus fica só aqui (não em LostFocus) pra evitar reentrância (erro 5),
'mesmo padrão usado no Entrada_Consulta_PorProdutos
AtualizarCriterioSec
If bCarregando Then Exit Sub   'ainda no Form_Load - SetFocus aqui dava erro 5
DoEvents
If txtCodBarra.Visible Then
   txtCodBarra.SetFocus
ElseIf cboDescricao.Visible Then
   cboDescricao.SetFocus
End If
End Sub

Private Sub cboCriterioSec_GotFocus()
SelectControl cboCriterioSec

Dim vTexto As String
vTexto = cboCriterioSec.Text

cboCriterioSec.Clear
PreencherConsulta

cboCriterioSec.Text = vTexto

moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
AtualizarCriterioSec
End Sub

Private Sub AtualizarCriterioPrinc()
If cboCriterioPrinc.Text = "TODOS" Then
   lblMes.Visible = False
   lblAno.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblInicio.Visible = False
   mskInicio.Visible = False
   cmdCalendario1.Visible = False
   lblFim.Visible = False
   mskFim.Visible = False
   cmdCalendario2.Visible = False
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
   lblMes.Visible = True
   lblAno.Visible = True
   cboMes.Visible = True
   cboAno.Visible = True
   lblInicio.Visible = False
   mskInicio.Visible = False
   cmdCalendario1.Visible = False
   lblFim.Visible = False
   mskFim.Visible = False
   cmdCalendario2.Visible = False
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
   lblMes.Visible = False
   lblAno.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblInicio.Visible = True
   mskInicio.Visible = True
   cmdCalendario1.Visible = True
   lblFim.Visible = True
   mskFim.Visible = True
   cmdCalendario2.Visible = True
End If
End Sub

Private Sub cboCriterioPrinc_Click()
AtualizarCriterioPrinc
If bCarregando Then Exit Sub   'ainda no Form_Load - SetFocus aqui dava erro 5
DoEvents
If cboMes.Visible Then
   cboMes.SetFocus
ElseIf mskInicio.Visible Then
   mskInicio.SetFocus
End If
End Sub

Private Sub cboCriterioPrinc_GotFocus()
SelectControl cboCriterioPrinc

Dim vTexto As String
vTexto = cboCriterioPrinc.Text

cboCriterioPrinc.Clear
PreencherCriterios

cboCriterioPrinc.Text = vTexto

moCombo.AttachTo cboCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_LostFocus()
AtualizarCriterioPrinc
End Sub

Private Sub cboMes_GotFocus()
cboMes.Clear

cboMes.AddItem "Janeiro"
cboMes.AddItem "Fevereiro"
cboMes.AddItem "Março"
cboMes.AddItem "Abril"
cboMes.AddItem "Maio"
cboMes.AddItem "Junho"
cboMes.AddItem "Julho"
cboMes.AddItem "Agosto"
cboMes.AddItem "Setembro"
cboMes.AddItem "Outubro"
cboMes.AddItem "Novembro"
cboMes.AddItem "Dezembro"

moCombo.AttachTo cboMes
End Sub

Private Sub cboMes_LostFocus()
cboAno.SetFocus
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 5
LastYear = iAno

For i = FirstYear To LastYear
   cboAno.AddItem i
Next

moCombo.AttachTo cboAno
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdBuscar_Click
End Sub

Private Sub cmdCalendario1_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty

Set fCal = New Calendario
fCal.Show vbModal

varData = fCal.DateSelected

Unload fCal
Set fCal = Nothing

If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub

mskInicio = Format(varData, "dd/mm/yy")
End Sub

Private Sub cmdCalendario2_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty

Set fCal = New Calendario
fCal.Show vbModal

varData = fCal.DateSelected

Unload fCal
Set fCal = Nothing

If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub

mskFim = Format(varData, "dd/mm/yy")
End Sub

Private Sub mskInicio_GotFocus()
SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub

Private Sub mskInicio_LostFocus()
If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
   mskInicio.Mask = ""
   mskInicio.Text = ""
   Exit Sub
Else
   If IsDate(mskInicio.Text) Then
      If mskFim.Visible Then mskFim.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskInicio.SetFocus
      SelectControl mskInicio
   End If
End If
End Sub

Private Sub mskFim_GotFocus()
SelectControl mskFim
End Sub

Private Sub mskFim_KeyPress(KeyAscii As Integer)
mskFim.Mask = "##/##/##"
End Sub

Private Sub mskFim_LostFocus()
If mskFim.Text = "" Or mskFim.Text = "__/__/__" Then
   mskFim.Mask = ""
   mskFim.Text = ""
   Exit Sub
Else
   If IsDate(mskFim.Text) Then
      cmdBuscar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFim.SetFocus
      SelectControl mskFim
   End If
End If
End Sub

Private Sub txtCodBarra_GotFocus()
SelectControl txtCodBarra
End Sub

Private Sub Form_Load()
bCarregando = True

Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

PreencherConsulta
cboCriterioSec.ListIndex = 1   'CÓD. BARRA
AtualizarCriterioSec

PreencherCriterios
cboCriterioPrinc.ListIndex = 0   'TODOS
AtualizarCriterioPrinc

cboForma.AddItem "TODOS"
cboForma.AddItem "CADASTRO"
cboForma.AddItem "ENTRADA"
cboForma.AddItem "AJUSTE"
cboForma.AddItem "CONTAGEM"
cboForma.AddItem "SAÍDA"
cboForma.AddItem "VENDA"
cboForma.AddItem "OFICINA"
cboForma.ListIndex = 0

Set moCombo = New cComboHelper
bCarregando = False
End Sub

'SetFocus no Form_Load da erro 5 (a janela ainda nao esta visivel
'nesse momento) - tem que ser no Activate, que roda depois que o
'form ja apareceu na tela
Private Sub Form_Activate()
If txtCodBarra.Visible Then
   txtCodBarra.SetFocus
ElseIf cboDescricao.Visible Then
   cboDescricao.SetFocus
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub
