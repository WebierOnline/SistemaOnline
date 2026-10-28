VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form VendasServicos_ConsultaAgrupada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENDAS - RESUMO DIÁRIO"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "VendasServicos_ConsultaAgrupada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   13215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
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
      Height          =   1815
      Left            =   60
      TabIndex        =   26
      Top             =   720
      Width           =   13095
      Begin VB.Frame frmFiltro1 
         Height          =   1515
         Left            =   5640
         TabIndex        =   32
         Top             =   180
         Width           =   7335
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   420
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1500
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   420
            Visible         =   0   'False
            Width           =   1155
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   315
            Left            =   2940
            TabIndex        =   8
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
            MICON           =   "VendasServicos_ConsultaAgrupada.frx":23D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   165
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   1515
            TabIndex        =   2
            Top             =   180
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1515
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   5475
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   420
            Width           =   2595
         End
         Begin VB.ComboBox cboFormaPgto 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   2595
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pgto"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1035
         End
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4740
      Picture         =   "VendasServicos_ConsultaAgrupada.frx":23EE
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   13065
      TabIndex        =   9
      Top             =   0
      Width           =   13095
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RESUMO DIÁRIO DE VENDAS"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   180
         Width           =   7515
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   240
         Picture         =   "VendasServicos_ConsultaAgrupada.frx":3426
         Stretch         =   -1  'True
         Top             =   0
         Width           =   900
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5175
      Left            =   60
      TabIndex        =   15
      Top             =   2580
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9128
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "DETALHAMENTO"
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
      MICON           =   "VendasServicos_ConsultaAgrupada.frx":9C6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirParcelas 
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "FINANCEIRO"
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
      MICON           =   "VendasServicos_ConsultaAgrupada.frx":9C88
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
      Left            =   4260
      TabIndex        =   5
      Top             =   7800
      Width           =   1635
      _ExtentX        =   2884
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
      MICON           =   "VendasServicos_ConsultaAgrupada.frx":9CA4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   8220
      Top             =   7800
      Width           =   2355
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serviços:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8460
      TabIndex        =   28
      Top             =   7860
      Width           =   630
   End
   Begin VB.Label lblTotalServicos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9180
      TabIndex        =   29
      Top             =   7860
      Width           =   1275
   End
   Begin VB.Label lblTotalProdutos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9180
      TabIndex        =   30
      Top             =   8160
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8460
      TabIndex        =   33
      Top             =   8160
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10920
      TabIndex        =   25
      Top             =   8460
      Width           =   690
   End
   Begin VB.Label lblTotalDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11640
      TabIndex        =   24
      Top             =   8460
      Width           =   1275
   End
   Begin VB.Label lblTotalAcresc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11640
      TabIndex        =   23
      Top             =   8760
      Width           =   1275
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acréscimo:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10800
      TabIndex        =   22
      Top             =   8760
      Width           =   780
   End
   Begin VB.Label lblSubtotalBruto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11640
      TabIndex        =   21
      Top             =   9060
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11220
      TabIndex        =   20
      Top             =   9060
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10980
      TabIndex        =   19
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11640
      TabIndex        =   18
      Top             =   8160
      Width           =   1275
   End
   Begin VB.Label lblQtda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11640
      TabIndex        =   14
      Top             =   7860
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant.:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11100
      TabIndex        =   13
      Top             =   7860
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   1635
      Left            =   10620
      Top             =   7800
      Width           =   2355
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   195
      Left            =   4620
      TabIndex        =   11
      Top             =   8340
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   8220
      Top             =   8580
      Width           =   2355
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balcão:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8460
      TabIndex        =   34
      Top             =   8640
      Width           =   510
   End
   Begin VB.Label lblTotalBalcao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9180
      TabIndex        =   35
      Top             =   8640
      Width           =   1275
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8460
      TabIndex        =   36
      Top             =   8940
      Width           =   525
   End
   Begin VB.Label lblTotalOficina 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9180
      TabIndex        =   37
      Top             =   8940
      Width           =   1275
   End
End
Attribute VB_Name = "VendasServicos_ConsultaAgrupada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String
Private mSoTodosTipo As Boolean  'True se cboTipo=TODOS - so entao a expansao Balcao/Oficina faz sentido (se ja filtrado por 1 tipo so, expandir mostraria a mesma linha)

Dim posX As Single

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Private Sub PreencherFormaPgto()
cboFormaPgto.Clear
cboFormaPgto.AddItem "TODOS"
cboFormaPgto.AddItem "À VISTA"
cboFormaPgto.AddItem "À PRAZO"
End Sub

Private Sub PreencherTipoConsulta()
cboTipo.Clear
cboTipo.AddItem "TODOS"
cboTipo.AddItem "BALCÃO"
cboTipo.AddItem "OFICINA"
End Sub


Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAno.AddItem i
Next

moCombo.AttachTo cboAno
End Sub

Private Sub cboAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdLocalizar_Click
End Sub


Private Sub cboFormaPgto_LostFocus()
'PreencherTipoPgto
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







Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

'Reaproveita a tela normal de consulta (nao-agrupada), ja filtrada pela
'linha clicada aqui - criterio PERIODO com a mesma data em inicio/fim (um
'unico dia), forma de pagamento e tipo (Balcao/Oficina, se uma linha filha
'ja expandida foi clicada; TODOS se foi a linha pai)
Private Sub cmdExibirProdutos_Click()
Dim lRow As Long, lRowPai As Long
Dim sDataFormatada As String, sFormaCrua As String, sFormaPgto As String, sTipo As String

lRow = Grid.Row
If lRow < 1 Then Exit Sub

If Grid.TextMatrix(lRow, 10) = "P" Then
    lRowPai = lRow
    sTipo = "TODOS"
ElseIf Grid.TextMatrix(lRow, 10) = "F" Then
    lRowPai = lRow
    Do While lRowPai > 1
        lRowPai = lRowPai - 1
        If Grid.TextMatrix(lRowPai, 10) = "P" Then Exit Do
    Loop
    sTipo = Trim(Grid.TextMatrix(lRow, 2))
Else
    Exit Sub
End If

sDataFormatada = Grid.TextMatrix(lRowPai, 1)
sFormaCrua = Grid.TextMatrix(lRowPai, 12)
If sFormaCrua = "À Vista" Then
    sFormaPgto = "À VISTA"
Else
    sFormaPgto = "À PRAZO"
End If

'gPularSetFocusConsulta (variavel GLOBAL, General.bas) precisa estar True
'ANTES do Load - o proprio Load ja dispara Form_Load, que ja seta
'cboCriterioPrinc.ListIndex=4 (MENSAL) e isso ja dispara
'cboCriterioPrinc_Change com seu proprio .SetFocus, ainda com a janela
'inexistente - se a flag fosse uma propriedade do form, so existiria DEPOIS
'do Load, tarde demais pra proteger esse primeiro Change
gPularSetFocusConsulta = True

Load VendasServicos_Consulta
With VendasServicos_Consulta
    'cboCriterioPrinc.ListIndex=2 (PERIODO) dispara cboCriterioPrinc_Change
    'de novo, com outro .SetFocus - flag continua True, continua protegido
    .cboCriterioPrinc.ListIndex = 2
    .mskInicio.Text = sDataFormatada
    .mskFim.Text = sDataFormatada
    .cboFormaPgto.Text = sFormaPgto
    .cboTipo.Text = sTipo
    .cmdLocalizar_Click
End With

'so agora a janela fica visivel de verdade - reseta a flag, daqui pra
'frente qualquer SetFocus do usuario interagindo normalmente volta a valer
gPularSetFocusConsulta = False
VendasServicos_Consulta.Show 1
End Sub

'Relatorio sempre com Balcao/Oficina detalhados (quando cboTipo=TODOS) -
'monta a mesma base de MostrarGridAgrupado, mas com um nivel extra (UNION
'ALL) agrupado tambem por tipo_pedido, pra cada dia+forma sair acompanhado
'das 2 sublinhas - independe de quais linhas o usuario expandiu no grid
Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset
Dim sSQLBase As String, sSQLPai As String, sSQLFilho As String, sSQLFinal As String
Dim varTipoConsulta As String, varFormaConsulta As String

If cboMes.ListIndex = -1 Or cboAno.ListIndex = -1 Then Exit Sub

CalcularFiltrosConsulta varTipoConsulta, varFormaConsulta

sSQLBase = "(SELECT p.COD_PEDIDO, p.DATA_COMPRA, p.TIPO_PAGAMENTO, p.tipo_pedido, p.SUBTOTAL, p.ValorDescReal, p.ValorAcrescReal, p.TOTAL, " & _
           "ISNULL((SELECT SUM(total) FROM pedidos_itens WHERE cod_pedido=p.COD_PEDIDO),0) AS vProdPedido, " & _
           "ISNULL((SELECT SUM(sv.total) FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=p.COD_PEDIDO),0) AS vServPedido " & _
           "FROM pedidos p " & _
           "WHERE (Month(p.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (Year(p.DATA_COMPRA) = " & cboAno & ") " & _
           "AND (p.CANCELADO = 0) AND (p.STATUS_PEDIDO = 1) AND (" & varTipoConsulta & ") AND (" & varFormaConsulta & ")) AS base"

sSQLPai = "SELECT DATA_COMPRA, TIPO_PAGAMENTO, '' AS var_tipo_exibicao, SUM(SUBTOTAL) AS SUBTOTAL, SUM(ValorDescReal) AS ValorDescReal, SUM(ValorAcrescReal) AS ValorAcrescReal, SUM(TOTAL) AS var_total, SUM(vProdPedido) AS vProd, SUM(vServPedido) AS vServ, 0 AS nivel FROM " & sSQLBase & " GROUP BY DATA_COMPRA, TIPO_PAGAMENTO"

If mSoTodosTipo Then
    sSQLFilho = "SELECT DATA_COMPRA, TIPO_PAGAMENTO, CASE tipo_pedido WHEN 'OFICINA' THEN '     OFICINA' ELSE '     BALCÃO' END AS var_tipo_exibicao, SUM(SUBTOTAL) AS SUBTOTAL, SUM(ValorDescReal) AS ValorDescReal, SUM(ValorAcrescReal) AS ValorAcrescReal, SUM(TOTAL) AS var_total, SUM(vProdPedido) AS vProd, SUM(vServPedido) AS vServ, 1 AS nivel FROM " & sSQLBase & " GROUP BY DATA_COMPRA, TIPO_PAGAMENTO, tipo_pedido"
    sSQLFinal = "SELECT * FROM (" & sSQLPai & " UNION ALL " & sSQLFilho & ") AS q ORDER BY DATA_COMPRA, TIPO_PAGAMENTO, nivel, var_tipo_exibicao"
Else
    sSQLFinal = sSQLPai & " ORDER BY DATA_COMPRA, TIPO_PAGAMENTO"
End If

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As ini

Set oIni = New ini
oIni.arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(sSQLFinal)

Set REL_VendasResumoDiario.Relatorio.Recordset = r

REL_VendasResumoDiario.dfQuant.Caption = lblQtda.Caption
REL_VendasResumoDiario.dfSubtotal.Caption = lblSubtotal.Caption
REL_VendasResumoDiario.dfDesconto.Caption = lblTotalDesc.Caption
REL_VendasResumoDiario.dfAcrescimo.Caption = lblTotalAcresc.Caption
REL_VendasResumoDiario.dfSubtotalBruto.Caption = lblSubtotalBruto.Caption
REL_VendasResumoDiario.dfProdutos.Caption = lblTotalProdutos.Caption
REL_VendasResumoDiario.dfServicos.Caption = lblTotalServicos.Caption
REL_VendasResumoDiario.dfBalcao.Caption = lblTotalBalcao.Caption
REL_VendasResumoDiario.dfOficina.Caption = lblTotalOficina.Caption
REL_VendasResumoDiario.dfTotal.Caption = lblSubtotalBruto.Caption

REL_VendasResumoDiario.lblTitulo.Caption = "RESUMO DIÁRIO DE VENDAS"

REL_VendasResumoDiario.rfForma.Caption = cboFormaPgto.Text

REL_VendasResumoDiario.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text

REL_VendasResumoDiario.Relatorio.Ativar
Unload REL_VendasResumoDiario

Me.Show 1
End Sub

Private Sub cmdLocalizar_Click()
MostrarGridAgrupado
End Sub



Private Sub Form_Load()
Set moCombo = New cComboHelper

Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherFormaPgto
cboFormaPgto.ListIndex = 0  'TODOS

'nesta versao agrupada nao ha 1 pedido por linha - cmdExibirProdutos
'(DETALHAMENTO) foi reaproveitado p/ abrir a consulta normal ja filtrada
'(ver cmdExibirProdutos_Click); cmdExibirParcelas ainda nao se aplica
cmdExibirParcelas.Enabled = False

'Tela simplificada: so Mes/Ano/Tipo/Forma Pgto importam aqui (a consulta e
'sempre mensal, agrupada por dia) - os demais criterios do form original
'(Vendedor/Cliente/Periodo/Codigo/Mecanico/Organizar por) e seus controles
'foram removidos do layout, nao existem mais neste form. Mes/Ano continuam
'Visible=False no design (so ficavam visiveis com MENSAL selecionado no
'form original) - precisam ser reativados aqui
lblMes.Visible = True
cboMes.Visible = True
lblAno.Visible = True
cboAno.Visible = True
lblMes.Top = 180
lblAno.Top = 180
cboMes.Top = 420
cboAno.Top = 420
cmdLocalizar.Top = 420
cmdLocalizar.Left = 2880

'mes/ano atual
cboMes_GotFocus
cboMes.ListIndex = Month(Date) - 1
cboAno_GotFocus
cboAno.ListIndex = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mSoTodosTipo And Grid.MouseRow >= 1 Then
      If Grid.TextMatrix(Grid.MouseRow, 10) = "P" Then
         Grid.ToolTipText = "Clique no + para ver BalcÃO/Oficina separados"
      Else
         Grid.ToolTipText = ""
      End If
   Else
      Grid.ToolTipText = ""
   End If
End Sub


'Calcula os 2 trechos de WHERE (filtro de tipo e de forma de pagamento) a
'partir dos combos - usado tanto por MostrarGridAgrupado quanto por
'cmdImprimir_Click, pra nao duplicar a logica
Private Sub CalcularFiltrosConsulta(ByRef varTipoConsulta As String, ByRef varFormaConsulta As String)
mSoTodosTipo = (cboTipo.Text = "TODOS")

If cboTipo.Text = "BALCÃO" Then
    varTipoConsulta = "p.tipo_pedido = 'VENDA'"
ElseIf cboTipo.Text = "OFICINA" Then
    varTipoConsulta = "p.tipo_pedido = 'OFICINA'"
Else
    varTipoConsulta = "p.tipo_pedido IN ('VENDA','OFICINA')"
End If

If cboFormaPgto.Text = "À VISTA" Then
    varFormaConsulta = "p.TIPO_PAGAMENTO = 'À Vista'"
ElseIf cboFormaPgto.Text = "À PRAZO" Then
    varFormaConsulta = "p.TIPO_PAGAMENTO = 'À Prazo'"
Else
    varFormaConsulta = "p.TIPO_PAGAMENTO IN ('À Vista','À Prazo')"
End If
End Sub

Private Sub MostrarGridAgrupado()
Dim rTabela As ADODB.Recordset
Dim sSQL As String
Dim varTipoConsulta As String
Dim varFormaConsulta As String

If cboMes.ListIndex = -1 Or cboAno.ListIndex = -1 Then Exit Sub

CalcularFiltrosConsulta varTipoConsulta, varFormaConsulta

'nivel PAI: agrupado por DATA_COMPRA + TIPO_PAGAMENTO (A Vista/A Prazo).
'produtos/servicos vem de subquery por pedido (igual ja era feito na tela
'nao-agrupada), depois somados no agregado externo
sSQL = "SELECT DATA_COMPRA, TIPO_PAGAMENTO, " & _
       "SUM(vProdPedido) AS vProdutos, SUM(vServPedido) AS vServicos, " & _
       "SUM(SUBTOTAL) AS vSubtotal, SUM(ValorDescReal) AS vDesc, SUM(ValorAcrescReal) AS vAcresc, SUM(TOTAL) AS vTotal " & _
       "FROM (" & _
          "SELECT p.COD_PEDIDO, p.DATA_COMPRA, p.TIPO_PAGAMENTO, p.SUBTOTAL, p.ValorDescReal, p.ValorAcrescReal, p.TOTAL, " & _
          "ISNULL((SELECT SUM(total) FROM pedidos_itens WHERE cod_pedido=p.COD_PEDIDO),0) AS vProdPedido, " & _
          "ISNULL((SELECT SUM(sv.total) FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=p.COD_PEDIDO),0) AS vServPedido " & _
          "FROM pedidos p " & _
          "WHERE (Month(p.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (Year(p.DATA_COMPRA) = " & cboAno & ") " & _
          "AND (p.CANCELADO = 0) AND (p.STATUS_PEDIDO = 1) AND (" & varTipoConsulta & ") AND (" & varFormaConsulta & ")" & _
       ") AS q " & _
       "GROUP BY DATA_COMPRA, TIPO_PAGAMENTO " & _
       "ORDER BY DATA_COMPRA, TIPO_PAGAMENTO"

Set rTabela = dbData.OpenRecordset(sSQL)
FormatarGridAgrupado rTabela
If Not rTabela Is Nothing Then
   If rTabela.State <> 0 Then rTabela.Close
End If
Set rTabela = Nothing

AtualizarTotaisBalcaoOficina varTipoConsulta, varFormaConsulta
End Sub

'lblTotalBalcao/lblTotalOficina nao vem do grid (as linhas PAI misturam
'Balcao+Oficina quando cboTipo=TODOS, e a divisao so aparece nas linhas
'filha quando o usuario expande manualmente - nem toda linha fica
'expandida) - calculado direto via query agregada por tipo_pedido, com os
'mesmos filtros de mes/ano/tipo/forma ja usados na consulta principal
Private Sub AtualizarTotaisBalcaoOficina(ByVal varTipoConsulta As String, ByVal varFormaConsulta As String)
Dim rTot As ADODB.Recordset
Dim sSQLTot As String
Dim dBalcao As Double, dOficina As Double

sSQLTot = "SELECT p.tipo_pedido, SUM(p.TOTAL) AS vTotal FROM pedidos p " & _
          "WHERE (Month(p.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (Year(p.DATA_COMPRA) = " & cboAno & ") " & _
          "AND (p.CANCELADO = 0) AND (p.STATUS_PEDIDO = 1) AND (" & varTipoConsulta & ") AND (" & varFormaConsulta & ") " & _
          "GROUP BY p.tipo_pedido"

Set rTot = dbData.OpenRecordset(sSQLTot)
dBalcao = 0
dOficina = 0
Do While Not rTot.EOF
   If rTot("tipo_pedido") = "VENDA" Then
      dBalcao = ValidateNull(rTot("vTotal"))
   ElseIf rTot("tipo_pedido") = "OFICINA" Then
      dOficina = ValidateNull(rTot("vTotal"))
   End If
   rTot.MoveNext
Loop
If rTot.State <> 0 Then rTot.Close
Set rTot = Nothing

lblTotalBalcao = Format(dBalcao, ocMONEY)
lblTotalOficina = Format(dOficina, ocMONEY)
End Sub

'Colunas do grid agrupado:
' 0=indicador(+/-)  1=DATA  2=TIPO(so preenchido nas linhas filha)
' 3=PRODUTOS  4=SERVICOS  5=SUBTOTAL  6=DESC.  7=ACRE.  8=TOTAL  9=FORMA
' 10=marcador oculto P(pai)/F(filha)  11=DATA_COMPRA crua(yyyy-mm-dd, p/re-consulta)
' 12=TIPO_PAGAMENTO cru (A Vista/A Prazo, p/re-consulta)
Private Sub FormatarGridAgrupado(rTabela As ADODB.Recordset)
Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear

      'Desliga o highlight visual de selecao do MSFlexGrid - sem isso, com
      'SelectionMode=flexSelectionByRow, o clique no "+"/"-" (Grid_MouseDown)
      'deixa uma selecao de linha ativa que sobrepoe com branco o
      'CellBackColor aplicado nas linhas-filha (mesmo bug e mesma correcao
      'definitiva ja usada em Parcelas_Consulta.frm)
      .Highlight = 0

      .Cols = 13
      .rows = 2

      .ColWidth(0) = 450
      .ColWidth(1) = 1000
      .ColWidth(2) = 1300
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1200
      .ColWidth(9) = 1000
      .ColWidth(10) = 0
      .ColWidth(11) = 0
      .ColWidth(12) = 0

      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "DATA"
      .TextMatrix(0, 2) = "TIPO"
      .TextMatrix(0, 3) = "PRODUTOS"
      .TextMatrix(0, 4) = "SERVIÇOS"
      .TextMatrix(0, 5) = "SUBTOTAL"
      .TextMatrix(0, 6) = "DESC."
      .TextMatrix(0, 7) = "ACRE."
      .TextMatrix(0, 8) = "TOTAL"
      .TextMatrix(0, 9) = "FORMA"
      .Redraw = False

      'colocar os cabecalhos em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i

      .ColAlignment(0) = 3
      .ColAlignment(1) = 3

      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 0) = IIf(mSoTodosTipo, "+", "")
            .TextMatrix(.rows - 1, 1) = Format(rTabela("DATA_COMPRA"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 2) = ""
            .TextMatrix(.rows - 1, 3) = Format(rTabela("vProdutos"), ocMONEY)
            .TextMatrix(.rows - 1, 4) = Format(rTabela("vServicos"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = Format(rTabela("vSubtotal"), ocMONEY)
            .TextMatrix(.rows - 1, 6) = Format(rTabela("vDesc"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = Format(rTabela("vAcresc"), ocMONEY)
            .TextMatrix(.rows - 1, 8) = Format(rTabela("vTotal"), ocMONEY)
            .TextMatrix(.rows - 1, 9) = rTabela("TIPO_PAGAMENTO")
            .TextMatrix(.rows - 1, 10) = "P"
            .TextMatrix(.rows - 1, 11) = Format(rTabela("DATA_COMPRA"), "yyyy-mm-dd")
            .TextMatrix(.rows - 1, 12) = rTabela("TIPO_PAGAMENTO")
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If

      .rows = .rows - 1

      'Zebra "data a data": linhas com a mesma DATA_COMPRA (coluna 11, cru
      'yyyy-mm-dd) ficam com a mesma cor de fundo - mesmo padrao do
      'Parcelas_Consulta.frm (FlexCores), so que agrupado por data em vez
      'de linha a linha, ja que aqui 1 data pode ter 2 linhas PAI (A Vista
      'e A Prazo). PRODUTOS/SERVICOS perderam o fundo amarelo fixo e agora
      'seguem a mesma cor mesclada das outras colunas
      FlexCoresPorData &HFFFFFF, &HE0E0E0

      'DATA, TOTAL, PRODUTOS e SERVICOS em negrito
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellFontBold = True
         .Col = 3
         .CellFontBold = True
         .Col = 4
         .CellFontBold = True
         .Col = 8
         .CellFontBold = True
      Next

      Grid.Redraw = True
   End With

   AtualizarTotaisAgrupado
picAguarde.Visible = False
End Sub

Private Function EImpar(ByVal iNum As Long) As Boolean
   EImpar = (iNum Mod 2)
End Function

'Zebra por grupo de DATA_COMPRA (coluna 11, valor cru yyyy-mm-dd): a cor so
'alterna quando a data muda de uma linha PAI pra outra - linhas PAI com a
'mesma data (ex: 1 em A Vista + 1 em A Prazo) ficam com a mesma cor. Linhas
'filha (ainda nao inseridas nesse ponto - FormatarGridAgrupado roda antes
'de qualquer expansao) nao entram nessa conta
Private Sub FlexCoresPorData(lCorPar As Long, lCorImpar As Long)
   Dim iLinha As Long
   Dim sDataAnterior As String
   Dim iGrupo As Integer
   Dim lCor As OLE_COLOR

   Grid.FillStyle = flexFillRepeat
   sDataAnterior = ""
   iGrupo = 0

   For iLinha = 1 To Grid.rows - 1
      If Grid.TextMatrix(iLinha, 11) <> sDataAnterior Then
         iGrupo = iGrupo + 1
         sDataAnterior = Grid.TextMatrix(iLinha, 11)
      End If

      If EImpar(iGrupo) Then
         lCor = lCorImpar
      Else
         lCor = lCorPar
      End If

      With Grid
         .Row = iLinha
         .Col = 1
         .ColSel = .Cols - 1
         .CellBackColor = lCor
      End With
   Next

   Grid.FillStyle = flexFillSingle
End Sub

'Expande/recolhe a divisao Balcao/Oficina de uma linha PAI - clique simples
'na coluna 0 (indicador +/-); so funciona quando cboTipo=TODOS (se ja
'filtrado por 1 tipo, a divisao seria identica a linha pai, sem sentido
'expandir). Usa MouseDown + MouseRow/MouseCol (nao Click + Row/Col) porque
'o grid esta com SelectionMode=flexSelectionByRow (selecao por linha
'inteira) - nesse modo Grid.Col nao reflete a coluna clicada de forma
'confiavel, mas MouseCol/MouseRow sempre refletem a posicao real do
'ponteiro, independente do SelectionMode
Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lRow As Long
lRow = Grid.MouseRow
If lRow < 1 Then Exit Sub
If Grid.MouseCol <> 0 Then Exit Sub
If Not mSoTodosTipo Then Exit Sub
If Grid.TextMatrix(lRow, 10) <> "P" Then Exit Sub

If Grid.TextMatrix(lRow, 0) = "-" Then
    RecolherLinha lRow
Else
    ExpandirLinha lRow
End If
End Sub

'Insere 2 linhas-filha (BALCAO/OFICINA) logo abaixo da linha PAI lRow, com
'os totais daquele dia+forma dividido por tipo - mesma formula agregada do
'nivel pai, so que filtrando pela DATA_COMPRA/TIPO_PAGAMENTO exatos (em vez
'de mes/ano inteiro) e agrupando por tipo_pedido em vez de data+forma
Private Sub ExpandirLinha(ByVal lRow As Long)
Dim rDet As ADODB.Recordset
Dim sSQL As String
Dim sDataCrua As String, sFormaCrua As String
Dim i As Long, c As Integer

sDataCrua = Grid.TextMatrix(lRow, 11)
sFormaCrua = Grid.TextMatrix(lRow, 12)

sSQL = "SELECT q.tipo_pedido, " & _
       "SUM(q.vProdPedido) AS vProdutos, SUM(q.vServPedido) AS vServicos, " & _
       "SUM(q.SUBTOTAL) AS vSubtotal, SUM(q.ValorDescReal) AS vDesc, SUM(q.ValorAcrescReal) AS vAcresc, SUM(q.TOTAL) AS vTotal " & _
       "FROM (" & _
          "SELECT p.tipo_pedido, p.SUBTOTAL, p.ValorDescReal, p.ValorAcrescReal, p.TOTAL, " & _
          "ISNULL((SELECT SUM(total) FROM pedidos_itens WHERE cod_pedido=p.COD_PEDIDO),0) AS vProdPedido, " & _
          "ISNULL((SELECT SUM(sv.total) FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=p.COD_PEDIDO),0) AS vServPedido " & _
          "FROM pedidos p " & _
          "WHERE (p.DATA_COMPRA = CONVERT(DATETIME, '" & sDataCrua & "', 120)) AND (p.TIPO_PAGAMENTO = '" & sFormaCrua & "') " & _
          "AND (p.CANCELADO = 0) AND (p.STATUS_PEDIDO = 1) AND (p.tipo_pedido IN ('VENDA','OFICINA'))" & _
       ") AS q " & _
       "GROUP BY q.tipo_pedido " & _
       "ORDER BY q.tipo_pedido"

Set rDet = dbData.OpenRecordset(sSQL)

Grid.Redraw = False
i = lRow
Do While Not rDet.EOF
    i = i + 1
    Grid.AddItem "", i
    Grid.TextMatrix(i, 0) = ""
    Grid.TextMatrix(i, 1) = ""
    If rDet("tipo_pedido") = "OFICINA" Then
        Grid.TextMatrix(i, 2) = "     OFICINA"
    Else
        Grid.TextMatrix(i, 2) = "     BALCÃO"
    End If
    Grid.TextMatrix(i, 3) = Format(rDet("vProdutos"), ocMONEY)
    Grid.TextMatrix(i, 4) = Format(rDet("vServicos"), ocMONEY)
    Grid.TextMatrix(i, 5) = Format(rDet("vSubtotal"), ocMONEY)
    Grid.TextMatrix(i, 6) = Format(rDet("vDesc"), ocMONEY)
    Grid.TextMatrix(i, 7) = Format(rDet("vAcresc"), ocMONEY)
    Grid.TextMatrix(i, 8) = Format(rDet("vTotal"), ocMONEY)
    Grid.TextMatrix(i, 9) = ""
    Grid.TextMatrix(i, 10) = "F"
    For c = 0 To 9
        Grid.Row = i
        Grid.Col = c
        Grid.CellBackColor = &HFFFFFF
        Grid.CellForeColor = RGB(0, 0, 139)
    Next c
    rDet.MoveNext
Loop
If rDet.State <> 0 Then rDet.Close
Set rDet = Nothing

Grid.TextMatrix(lRow, 0) = "-"
Grid.Redraw = True
End Sub

'Remove as linhas-filha (marcador F) logo abaixo da linha PAI lRow
Private Sub RecolherLinha(ByVal lRow As Long)
Grid.Redraw = False
Do While lRow + 1 <= Grid.rows - 1
    If Grid.TextMatrix(lRow + 1, 10) = "F" Then
        Grid.RemoveItem lRow + 1
    Else
        Exit Do
    End If
Loop
Grid.TextMatrix(lRow, 0) = "+"
Grid.Redraw = True
End Sub

'Soma uma coluna do grid contando SO as linhas PAI (marcador=P) - as linhas
'filha (Balcao/Oficina expandidas) sao so uma quebra visual do MESMO valor
'ja contado na linha pai, somar as duas duplicaria o total
Public Function SomaGridSeMarcador(var_Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency

   Valor = 0

   For i = 1 To var_Grid.rows - 1
      If var_Grid.TextMatrix(i, 10) = "P" Then
         If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
            Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
         End If
      End If
   Next

   SomaGridSeMarcador = Valor
End Function

Private Sub AtualizarTotaisAgrupado()
Dim i As Integer, qtdGrupos As Integer
qtdGrupos = 0
For i = 1 To Grid.rows - 1
   If Grid.TextMatrix(i, 10) = "P" Then qtdGrupos = qtdGrupos + 1
Next

lblQtda = Format(qtdGrupos, "0")
lblTotalProdutos = Format(SomaGridSeMarcador(Grid, 3), ocMONEY)
lblTotalServicos = Format(SomaGridSeMarcador(Grid, 4), ocMONEY)
lblSubtotal = Format(SomaGridSeMarcador(Grid, 5), ocMONEY)
lblTotalDesc = Format(SomaGridSeMarcador(Grid, 6), ocMONEY)
lblTotalAcresc = Format(SomaGridSeMarcador(Grid, 7), ocMONEY)
lblSubtotalBruto = Format(SomaGridSeMarcador(Grid, 8), ocMONEY)
End Sub

