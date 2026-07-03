VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Vendas_Consulta_PorProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA PRODUTOS E SERVIÇOS"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   Icon            =   "Vendas_Consulta_PorProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13395
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   8700
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   450
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
      MICON           =   "Vendas_Consulta_PorProdutos.frx":23D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   60
      ScaleHeight     =   1725
      ScaleWidth      =   13245
      TabIndex        =   8
      ToolTipText     =   "Imprimir"
      Top             =   660
      Width           =   13275
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1515
         Left            =   120
         TabIndex        =   27
         Top             =   60
         Width           =   4515
         Begin VB.ComboBox cboIndice 
            Height          =   315
            Left            =   2280
            TabIndex        =   31
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cboCriterioSec 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   2115
         End
         Begin VB.ComboBox cboCriterioPrinc 
            Height          =   315
            Left            =   2280
            TabIndex        =   30
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   2115
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Organizar por:"
            Height          =   195
            Left            =   2280
            TabIndex        =   35
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Criterio"
            Height          =   195
            Left            =   2280
            TabIndex        =   34
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tipo"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Consultar por:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Escolha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1515
         Left            =   4680
         TabIndex        =   9
         Top             =   60
         Width           =   8475
         Begin VB.TextBox txtCodProduto 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5760
            TabIndex        =   38
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ChamaleonBtn.chameleonButton cmdCalendario1 
            Height          =   315
            Left            =   1140
            TabIndex        =   25
            Tag             =   "Calendario"
            Top             =   1080
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
            MICON           =   "Vendas_Consulta_PorProdutos.frx":23EE
            PICN            =   "Vendas_Consulta_PorProdutos.frx":240A
            PICH            =   "Vendas_Consulta_PorProdutos.frx":475D
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
            Left            =   2880
            TabIndex        =   26
            Tag             =   "Calendario"
            Top             =   1080
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
            MICON           =   "Vendas_Consulta_PorProdutos.frx":6AB0
            PICN            =   "Vendas_Consulta_PorProdutos.frx":6ACC
            PICH            =   "Vendas_Consulta_PorProdutos.frx":8E1F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   180
            TabIndex        =   16
            Top             =   480
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.ComboBox cboDescricao 
            Height          =   315
            Left            =   180
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   8175
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   180
            TabIndex        =   17
            Top             =   1080
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFim 
            Height          =   315
            Left            =   1920
            TabIndex        =   18
            Top             =   1080
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yy"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   495
            Left            =   6900
            TabIndex        =   37
            Top             =   960
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
            MICON           =   "Vendas_Consulta_PorProdutos.frx":B172
            PICN            =   "Vendas_Consulta_PorProdutos.frx":B18E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   180
            TabIndex        =   14
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   1080
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblDescricao 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   840
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   1560
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lblInicio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data inicial:"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   840
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lblFim 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data final:"
            Height          =   195
            Left            =   1920
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblAte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "até"
            Height          =   195
            Left            =   1560
            TabIndex        =   19
            Top             =   1140
            Visible         =   0   'False
            Width           =   225
         End
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4380
      Picture         =   "Vendas_Consulta_PorProdutos.frx":BA68
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   60
      ScaleHeight     =   525
      ScaleWidth      =   13245
      TabIndex        =   0
      Top             =   60
      Width           =   13275
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE PRODUTOS E SERVIÇOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1380
         TabIndex        =   1
         Top             =   120
         Width           =   4740
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   900
         Picture         =   "Vendas_Consulta_PorProdutos.frx":CAA0
         Top             =   40
         Width           =   450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6195
      Left            =   60
      TabIndex        =   11
      Top             =   2460
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   10927
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirParcelas 
      Height          =   255
      Left            =   2340
      TabIndex        =   12
      Top             =   8700
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   450
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
      MICON           =   "Vendas_Consulta_PorProdutos.frx":D3C7
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
      Height          =   255
      Left            =   4620
      TabIndex        =   36
      Top             =   8700
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   450
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
      MICON           =   "Vendas_Consulta_PorProdutos.frx":D3E3
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
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
      Left            =   10860
      TabIndex        =   7
      Top             =   9180
      Width           =   675
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   11580
      TabIndex        =   6
      Top             =   9180
      Width           =   1635
   End
   Begin VB.Label lblQtda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   11580
      TabIndex        =   5
      Top             =   8820
      Width           =   1635
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANT.:"
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
      Left            =   10740
      TabIndex        =   4
      Top             =   8820
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   915
      Left            =   10620
      Top             =   8700
      Width           =   2715
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Vendas_Consulta_PorProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private printSQL As String

Dim posX As Single

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long
Private Sub FormatarGrid_ProdutosLucros(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 5
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 6760
      .ColWidth(2) = 1000
      .ColWidth(3) = 800
      .ColWidth(4) = 1000
      
      .TextMatrix(0, 1) = "DESCRIÇÃO"
      .TextMatrix(0, 2) = "PREÇO"
      .TextMatrix(0, 3) = "QTDE"
      .TextMatrix(0, 4) = "TOTAL"
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      .ColAlignment(1) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("descricao")
            .TextMatrix(.rows - 1, 2) = Format$(rTabela("preco"), ocMONEY)
            .TextMatrix(.rows - 1, 3) = rTabela("var_qtde")
            .TextMatrix(.rows - 1, 4) = Format$(rTabela("var_total"), ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Redraw = True
      .rows = .rows - 1
   End With
   
   lblQtda.Caption = Format(SomaGrid(Grid, 3), ocPESO)
   lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 13
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 750
      .ColWidth(2) = 900
      .ColWidth(3) = 0
      .ColWidth(4) = 1200
      .ColWidth(5) = 5300
      .ColWidth(6) = 800
      .ColWidth(7) = 800
      .ColWidth(8) = 850
      .ColWidth(9) = 700
      .ColWidth(10) = 800
      .ColWidth(11) = 0
      .ColWidth(12) = 0
      
      .TextMatrix(0, 1) = "OS"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = ""
      .TextMatrix(0, 4) = "CÓD. SERV."
      .TextMatrix(0, 5) = "DESCRIÇÃO"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "QTDE"
      .TextMatrix(0, 8) = "SUBTOTAL"
      .TextMatrix(0, 9) = "DESC."
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = ""
      .TextMatrix(0, 12) = "CÓD.PROD."
      
      .Redraw = False
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 1
      
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("varCodPed"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("varData"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = ""
            .TextMatrix(.rows - 1, 4) = rTabela("varCodServ")
            .TextMatrix(.rows - 1, 5) = rTabela("varNome")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("varValor"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = rTabela("varQuant")
            .TextMatrix(.rows - 1, 8) = Format(rTabela("varSubtotal"), ocMONEY)
            .TextMatrix(.rows - 1, 9) = Format(rTabela("varDesc"), ocMONEY)
            .TextMatrix(.rows - 1, 10) = Format(rTabela("varTotal"), ocMONEY)
            .TextMatrix(.rows - 1, 11) = rTabela("var_CodOS")
            .TextMatrix(.rows - 1, 12) = ""
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblQtda.Caption = SomaGrid(Grid, 7)
   lblTotal.Caption = Format(SomaGrid(Grid, 10), ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub FormatarGrid_ProdDetalhado(rTabela As ADODB.Recordset)
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 13
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 700
      .ColWidth(2) = 800
      .ColWidth(3) = 0
      .ColWidth(4) = 1200
      .ColWidth(5) = 5300
      .ColWidth(6) = 800
      .ColWidth(7) = 700
      .ColWidth(8) = 850
      .ColWidth(9) = 700
      .ColWidth(10) = 800
      .ColWidth(11) = 0
      .ColWidth(12) = 900
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "CÓD.PROD."
      .TextMatrix(0, 4) = "CÓD. BARRA"
      .TextMatrix(0, 5) = "DESCRIÇÃO"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "QTDE"
      .TextMatrix(0, 8) = "SUBTOTAL"
      .TextMatrix(0, 9) = "DESC."
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = "COD_OS"
      .TextMatrix(0, 12) = "CÓD.PROD."
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      .ColAlignment(1) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
      
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("varcodped"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("varData"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = rTabela("varcodprod")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("varCodBarra"))
            '.TextMatrix(.Rows - 1, 5) = rTabela("vardesc")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("varvalor"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = rTabela("varquant")
            .TextMatrix(.rows - 1, 8) = Format(rTabela("varsubtotal"), ocMONEY)
            .TextMatrix(.rows - 1, 9) = Format(rTabela("vardesc"), ocMONEY)
            .TextMatrix(.rows - 1, 10) = Format(rTabela("vartotal"), ocMONEY)
            .TextMatrix(.rows - 1, 11) = rTabela("var_codos")
            .TextMatrix(.rows - 1, 12) = rTabela("varcodprod")
            
            If tipoEmpresa = 4 Then
            .TextMatrix(.rows - 1, 5) = rTabela("varNome") & " /  " & rTabela("vartam") & " / " & rTabela("varfab") & " /  " & rTabela("varref")
            Else
            .TextMatrix(.rows - 1, 5) = rTabela("varNome") & " /  " & ValidateNull(rTabela("varfab")) & " /  " & rTabela("varRef")
            End If
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      'For i = 1 To .Rows - 1
      '   .Row = i
      '   .Col = 3
      '   .CellForeColor = &HC0&
      '   .CellFontBold = True
      'Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblQtda.Caption = SomaGrid(Grid, 7)
   lblTotal.Caption = Format(SomaGrid(Grid, 10), ocMONEY)
   'lblEntrada.Caption = Format(0, ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub FormatarGrid_Produtos(rTabela As ADODB.Recordset)
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 5
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 6660
      .ColWidth(2) = 1000
      .ColWidth(3) = 900
      .ColWidth(4) = 1000
      
      .TextMatrix(0, 1) = "DESCRIÇÃO"
      .TextMatrix(0, 2) = "PREÇO"
      .TextMatrix(0, 3) = "QTDE"
      .TextMatrix(0, 4) = "TOTAL"
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      .ColAlignment(1) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      If Not rTabela Is Nothing Then
      
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 0) = rTabela("cod_produto")
            
            If tipoEmpresa = 4 Then
            .TextMatrix(.rows - 1, 1) = rTabela("var_desc") & " /  " & rTabela("var_tam") & " / " & rTabela("var_fab") & " /  " & rTabela("ref")
            Else
            .TextMatrix(.rows - 1, 1) = rTabela("var_desc") & " /  " & ValidateNull(rTabela("var_fab"))
            End If
            
            .TextMatrix(.rows - 1, 2) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.rows - 1, 3) = rTabela("var_qtde")
            .TextMatrix(.rows - 1, 4) = Format(rTabela("var_total"), ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      .Redraw = True
   End With
   
   lblQtda.Caption = Format(SomaGrid(Grid, 3), ocPESO)
   lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub Limpar_Grid_Venda()
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1220
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "TIPO"
      .TextMatrix(0, 7) = "TIPO"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   lblQtda.Caption = Format(0, ocMONEY)
   lblTotal.Caption = Format(0, ocMONEY)
picAguarde.Visible = False
End Sub

Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
mskFim.Mask = ""
mskFim.Text = ""
mskInicio.Mask = ""
mskInicio.Text = ""
End Sub

Private Sub PreencherCriterio()
cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "MENSAL"
End Sub

Private Sub PreencherCriterioSec()
cboCriterioSec.AddItem "DESCRIÇÃO"
cboCriterioSec.AddItem "CÓD. BARRA"
cboCriterioSec.AddItem "REFERÊNCIA"
cboCriterioSec.AddItem "FABRICANTE"
End Sub

Private Sub PreencherIndice()
cboIndice.AddItem "QUANT."
cboIndice.AddItem "PRODUTO"
cboIndice.AddItem "DATA"
cboIndice.AddItem "PEDIDO"
End Sub

Private Sub PreencherTipoConsulta()
cboTipo.AddItem "POR PRODUTOS"
cboTipo.AddItem "POR SERVIÇOS"
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

Private Sub cboCriterioPrinc_Click()
cboCriterioPrinc_LostFocus
End Sub

Private Sub cboCriterioPrinc_GotFocus()
cboCriterioPrinc.Clear
   
If cboTipo.Text = "POR PRODUTOS" Then
   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"
   cboCriterioPrinc.AddItem "MENSAL"
   cboCriterioPrinc.AddItem "PERÍODO"
   cboCriterioPrinc.AddItem "DATA"
   If cboCriterioSec.Text = "TODOS" Then
      cboCriterioPrinc.AddItem "PRODUTO/MENSAL"
      cboCriterioPrinc.AddItem "PRODUTO/PERÍODO"
   End If
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"
   cboCriterioPrinc.AddItem "MENSAL"
   cboCriterioPrinc.AddItem "PERÍODO"
   cboCriterioPrinc.AddItem "SERVIÇOS"
   If cboCriterioSec.Text = "TODOS" Then
      cboCriterioPrinc.AddItem "SERVIÇOS/MENSAL"
      cboCriterioPrinc.AddItem "SERVIÇOS/PERÍODO"
   End If
End If
   
moCombo.AttachTo cboCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_LostFocus()
If cboCriterioPrinc.Text = "TODOS" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
    lblInicio.Visible = True
    lblInicio.Caption = "Inicio"
    mskInicio.Visible = True
    lblFim.Visible = True
    mskFim.Visible = True
    lblAte.Visible = True
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = True
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
ElseIf cboCriterioPrinc.Text = "DATA" Then
    lblInicio.Visible = True
    lblInicio.Caption = "Data"
    mskInicio.Visible = True
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
ElseIf cboCriterioPrinc.Text = "PRODUTO/MENSAL" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
    lblDescricao.Caption = "Produto"
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
    LimparObjetos_Consulta
    Exit Sub
ElseIf cboCriterioPrinc.Text = "PRODUTO/PERÍODO" Then
    lblInicio.Visible = True
    lblInicio.Caption = "Inicio"
    mskInicio.Visible = True
    lblFim.Visible = True
    mskFim.Visible = True
    lblAte.Visible = True
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = True
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblDescricao.Caption = "Produto"
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
    LimparObjetos_Consulta
    Exit Sub
ElseIf cboCriterioPrinc.Text = "SERVIÇOS" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblDescricao.Caption = "Serviço"
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
    cboDescricao.Text = ""
    LimparObjetos_Consulta
    Exit Sub
ElseIf cboCriterioPrinc.Text = "SERVIÇOS/MENSAL" Then
    lblInicio.Visible = False
    mskInicio.Visible = False
    lblFim.Visible = False
    mskFim.Visible = False
    lblAte.Visible = False
    cmdCalendario1.Visible = False
    cmdCalendario2.Visible = False
    lblMes.Visible = True
    cboMes.Visible = True
    lblAno.Visible = True
    cboAno.Visible = True
    lblDescricao.Caption = "Serviço"
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
    cboDescricao.Text = ""
    LimparObjetos_Consulta
    Exit Sub
ElseIf cboCriterioPrinc.Text = "SERVIÇOS/PERÍODO" Then
    lblInicio.Visible = True
    lblInicio.Caption = "Inicio"
    mskInicio.Visible = True
    lblFim.Visible = True
    mskFim.Visible = True
    lblAte.Visible = True
    cmdCalendario1.Visible = True
    cmdCalendario2.Visible = True
    lblMes.Visible = False
    cboMes.Visible = False
    lblAno.Visible = False
    cboAno.Visible = False
    lblDescricao.Caption = "Serviço"
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
    cboDescricao.Text = ""
    LimparObjetos_Consulta
    Exit Sub
End If

If cboCriterioSec.Text = "DESCRIÇÃO" Or cboCriterioSec.Text = "REFERÊNCIA" Or cboCriterioSec.Text = "FABRICANTE" Then
    If cboCriterioSec.Text = "DESCRIÇÃO" Then
        lblDescricao.Caption = "Descrição"
    ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
        lblDescricao.Caption = "Referência"
    ElseIf cboCriterioSec.Text = "FABRICANTE" Then
        lblDescricao.Caption = "Fabricante"
    End If
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
    lblDescricao.Caption = "Cód. Barra"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
ElseIf cboCriterioSec.Text = "CÓD. OS" Then
    lblDescricao.Caption = "Cód. OS"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
ElseIf cboCriterioSec.Text = "TODOS" Then
    lblDescricao.Visible = False
    cboDescricao.Visible = False
    txtCodBarra.Visible = False
Else
End If


LimparObjetos_Consulta
End Sub


Private Sub cboCriterioSec_Click()
cboCriterioSec_LostFocus
End Sub

Private Sub cboCriterioSec_GotFocus()
cboCriterioSec.Clear

If cboTipo.Text = "POR SERVIÇOS" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "DESCRIÇÃO"
   cboCriterioSec.AddItem "CÓD. OS"
Else
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "DESCRIÇÃO"
   cboCriterioSec.AddItem "CÓD. BARRA"
   cboCriterioSec.AddItem "REFERÊNCIA"
   cboCriterioSec.AddItem "FABRICANTE"
   cboCriterioSec.AddItem "CATEGORIA"
End If

moCombo.AttachTo cboCriterioSec
End Sub

Private Sub cboCriterioSec_LostFocus()
If cboCriterioSec.Text = "DESCRIÇÃO" Or cboCriterioSec.Text = "REFERÊNCIA" Or cboCriterioSec.Text = "FABRICANTE" Or cboCriterioSec.Text = "CATEGORIA" Then
    If cboCriterioSec.Text = "DESCRIÇÃO" Then
        lblDescricao.Caption = "Descrição"
    ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
        lblDescricao.Caption = "Referência"
    ElseIf cboCriterioSec.Text = "FABRICANTE" Then
        lblDescricao.Caption = "Fabricante"
    ElseIf cboCriterioSec.Text = "CATEGORIA" Then
        lblDescricao.Caption = "Categoria"
    End If
    lblDescricao.Visible = True
    cboDescricao.Visible = True
    txtCodBarra.Visible = False
ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
    lblDescricao.Caption = "Cód. Barra"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
ElseIf cboCriterioSec.Text = "CÓD. OS" Then
    lblDescricao.Caption = "Cód. OS"
    lblDescricao.Visible = True
    cboDescricao.Visible = False
    txtCodBarra.Visible = True
ElseIf cboCriterioSec.Text = "TODOS" Then
    lblDescricao.Visible = False
    cboDescricao.Visible = False
    txtCodBarra.Visible = False
Else
End If

cboCriterioPrinc.Clear
If cboTipo.Text = "POR PRODUTOS" Then
   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"
   cboCriterioPrinc.AddItem "MENSAL"
   cboCriterioPrinc.AddItem "PERÍODO"
   cboCriterioPrinc.AddItem "DATA"
   If cboCriterioSec.Text = "TODOS" Then
      cboCriterioPrinc.AddItem "PRODUTO/MENSAL"
      cboCriterioPrinc.AddItem "PRODUTO/PERÍODO"
   End If
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"
   cboCriterioPrinc.AddItem "MENSAL"
   cboCriterioPrinc.AddItem "PERÍODO"
   cboCriterioPrinc.AddItem "SERVIÇOS"
   If cboCriterioSec.Text = "TODOS" Then
      cboCriterioPrinc.AddItem "SERVIÇOS/MENSAL"
      cboCriterioPrinc.AddItem "SERVIÇOS/PERÍODO"
   End If
End If
cboDescricao.Text = ""
cboCriterioPrinc.ListIndex = 0
cboCriterioPrinc_LostFocus
End Sub


Private Sub cboDescricao_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboDescricao.Clear
   
If cboTipo.Text = "POR SERVIÇOS" Then
   sSQL = "SELECT SERVICO, CODIGO FROM OS_Servicos ORDER BY SERVICO;"
   Set r = dbData.OpenRecordset(sSQL)
   Do While Not r.EOF
      cboDescricao.AddItem r("SERVICO")
      cboDescricao.ItemData(cboDescricao.NewIndex) = r("CODIGO")
      r.MoveNext
   Loop
   If r.State <> 0 Then r.Close
   Set r = Nothing
   moCombo.AttachTo cboDescricao
   Exit Sub
End If

If cboCriterioPrinc.Text = "PRODUTO/MENSAL" Or cboCriterioPrinc.Text = "PRODUTO/PERÍODO" Then
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   If r.State <> 0 Then r.Close
   Set r = Nothing
   moCombo.AttachTo cboDescricao
   Exit Sub
End If

If cboCriterioSec.Text = "DESCRIÇÃO" Then
   sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
   sSQL = "SELECT DISTINCT REF FROM produtos ORDER BY REF;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("REF"))
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "FABRICANTE" Then
   sSQL = "SELECT DISTINCT FABRICANTE FROM produtos ORDER BY FABRICANTE;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("FABRICANTE"))
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   sSQL = "SELECT DISTINCT CATEGORIA FROM produtos ORDER BY CATEGORIA;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("CATEGORIA"))
      r.MoveNext
   Loop
Else
   Exit Sub
End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboDescricao
End Sub


Private Sub cboDescricao_LostFocus()
On Error GoTo TrataErro
If cboDescricao.Text = "" Then txtCodProduto.Text = "": Exit Sub

txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboIndice_GotFocus()
cboIndice.Clear
cboIndice.AddItem "QUANT."
cboIndice.AddItem "PRODUTO"
cboIndice.AddItem "DATA"
cboIndice.AddItem "PEDIDO"
moCombo.AttachTo cboIndice
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

Private Sub cboTipo_Change()
If cboTipo.Text = "POR PRODUTOS" Then
'cmdExibirParcelas.Visible = False
   cmdExibirPedidos.Visible = True
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
'cmdExibirParcelas.Visible = False
   cmdExibirPedidos.Visible = True
Else
   Exit Sub
End If

' Recarrega cboCriterioSec e seleciona TODOS
cboCriterioSec.Clear
If cboTipo.Text = "POR SERVIÇOS" Then
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "DESCRIÇÃO"
   cboCriterioSec.AddItem "CÓD. OS"
Else
   cboCriterioSec.AddItem "TODOS"
   cboCriterioSec.AddItem "DESCRIÇÃO"
   cboCriterioSec.AddItem "CÓD. BARRA"
   cboCriterioSec.AddItem "REFERÊNCIA"
   cboCriterioSec.AddItem "FABRICANTE"
   cboCriterioSec.AddItem "CATEGORIA"
End If
cboCriterioSec.ListIndex = 0

' Recarrega cboCriterioPrinc sem TODOS (cboCriterioSec = TODOS)
cboCriterioPrinc.Clear
If cboTipo.Text = "POR PRODUTOS" Then
   cboCriterioPrinc.AddItem "MENSAL"
   cboCriterioPrinc.AddItem "PERÍODO"
   cboCriterioPrinc.AddItem "DATA"
   cboCriterioPrinc.AddItem "PRODUTO/MENSAL"
   cboCriterioPrinc.AddItem "PRODUTO/PERÍODO"
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   cboCriterioPrinc.AddItem "MENSAL"
   cboCriterioPrinc.AddItem "PERÍODO"
   cboCriterioPrinc.AddItem "SERVIÇOS"
   cboCriterioPrinc.AddItem "SERVIÇOS/MENSAL"
   cboCriterioPrinc.AddItem "SERVIÇOS/PERÍODO"
End If
cboCriterioPrinc.ListIndex = 0

cboCriterioSec_LostFocus
cboCriterioPrinc_LostFocus
End Sub

Private Sub cboTipo_Click()
cboTipo_Change
End Sub

Private Sub cboTipo_GotFocus()
cboTipo.Clear
cboTipo.AddItem "POR PRODUTOS"
cboTipo.AddItem "POR SERVIÇOS"
moCombo.AttachTo cboTipo
End Sub

Private Sub cmdCalendario1_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCalendario2_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskFim = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   Dim lPedido As Long
   If cboTipo.Text = "POR SERVIÇOS" Then
      If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub
      lPedido = CLng(Grid.TextMatrix(Grid.Row, 11))
      If lPedido = 0 Then Exit Sub
   ElseIf IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then
      lPedido = CLng(Grid.TextMatrix(Grid.Row, 1))
   Else
      Exit Sub
   End If
   Vendas_Consulta_Geral_Parcelas.loadInformacoes lPedido
   Vendas_Consulta_Geral_Parcelas.Show 1
End Sub

Private Sub cmdExibirPedidos_Click()
''If cboTipo.Text = "POR PRODUTOS" Then
''   If Grid.Col = 0 Then Exit Sub
''   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
''      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
''      Vendas_Consulta_Pedidos.loadPedidos Grid.TextMatrix(Grid.Row, 1)
''      Vendas_Consulta_Pedidos.Show 1
''   End If
''End If

'If Grid.Col = 0 Then Exit Sub
'If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
'   If Grid.Col = 1 Then
'      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
'      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 7)
'      Parcelas_Consulta_Produtos.Show 1
'   End If
'End If


If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub

If cboTipo.Text = "POR SERVIÇOS" Then
   ' col1 = COD_OS, col11 = COD_PEDIDO (gravado pelo FormatarGrid_Servicos)
   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub
   If CLng(Grid.TextMatrix(Grid.Row, 11)) = 0 Then Exit Sub
   Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 11)), "OS"
Else
   ' POR PRODUTOS: col1 = cod_pedido, col11 = COD_OS (0 = sem OS)
   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then Exit Sub
   If Grid.TextMatrix(Grid.Row, 11) <> "0" Then
      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), "OS"
   Else
      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), "VENDA"
   End If
End If

Parcelas_Consulta_Produtos.Show 1
End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset

Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)
Set REL_Cons_Venda_Prod.Relatorio.Recordset = r

If cboTipo.Text = "POR PRODUTOS" Then

    If cboCriterioPrinc.Text = "TODOS" Then
        REL_Cons_Venda_Prod.rfCons1.Caption = "TODOS"
        REL_Cons_Venda_Prod.rfCons3.Caption = ""
    ElseIf cboCriterioPrinc.Text = "MENSAL" Then
        REL_Cons_Venda_Prod.rfCons1.Caption = "MENSAL"
        REL_Cons_Venda_Prod.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
    ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
        REL_Cons_Venda_Prod.rfCons1.Caption = "PERÍODO"
        REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
    ElseIf cboCriterioPrinc.Text = "DATA" Then
        REL_Cons_Venda_Prod.rfCons1.Caption = "DATA"
        REL_Cons_Venda_Prod.rfCons3.Caption = "Data = " & mskInicio.Text
    ElseIf cboCriterioPrinc.Text = "PRODUTO/MENSAL" Then
        REL_Cons_Venda_Prod.rfCons1.Caption = "PRODUTO/MENSAL"
        REL_Cons_Venda_Prod.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
    ElseIf cboCriterioPrinc.Text = "PRODUTO/PERÍODO" Then
        REL_Cons_Venda_Prod.rfCons1.Caption = "PRODUTO/PERÍODO"
        REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
    End If

    If cboCriterioPrinc.Text = "PRODUTO/MENSAL" Or cboCriterioPrinc.Text = "PRODUTO/PERÍODO" Then
        REL_Cons_Venda_Prod.rfCons2.Caption = "PRODUTO = " & cboDescricao.Text
    ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
        REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRIÇÃO = " & cboDescricao.Text & ""
    ElseIf cboCriterioSec.Text = "CÓD. BARRA" Then
        REL_Cons_Venda_Prod.rfCons2.Caption = "CÓD. BARRA = " & txtCodBarra.Text & ""
    ElseIf cboCriterioSec.Text = "REFERÊNCIA" Then
        REL_Cons_Venda_Prod.rfCons2.Caption = "REFERÊNCIA = " & cboDescricao.Text & ""
    ElseIf cboCriterioSec.Text = "FABRICANTE" Then
        REL_Cons_Venda_Prod.rfCons2.Caption = "FABRICANTE = " & cboDescricao.Text & ""
    ElseIf cboCriterioSec.Text = "CATEGORIA" Then
        REL_Cons_Venda_Prod.rfCons2.Caption = "CATEGORIA = " & cboDescricao.Text & ""
    End If

ElseIf cboTipo.Text = "POR SERVIÇOS" Then

    REL_Cons_Venda_Prod.rfCons1.Caption = cboCriterioPrinc.Text
    REL_Cons_Venda_Prod.rfCons2.Caption = ""
    REL_Cons_Venda_Prod.rfCons3.Caption = ""

    If cboCriterioPrinc.Text = "TODOS" Then
        If cboCriterioSec.Text = "DESCRIÇÃO" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRIÇÃO = " & cboDescricao.Text
        ElseIf cboCriterioSec.Text = "CÓD. OS" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "CÓD. OS = " & txtCodBarra.Text
        End If

    ElseIf cboCriterioPrinc.Text = "MENSAL" Then
        If cboCriterioSec.Text = "DESCRIÇÃO" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRIÇÃO = " & cboDescricao.Text
            REL_Cons_Venda_Prod.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
        ElseIf cboCriterioSec.Text = "CÓD. OS" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "CÓD. OS = " & txtCodBarra.Text
            REL_Cons_Venda_Prod.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
        Else
            REL_Cons_Venda_Prod.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
        End If

    ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
        If cboCriterioSec.Text = "DESCRIÇÃO" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRIÇÃO = " & cboDescricao.Text
            REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
        ElseIf cboCriterioSec.Text = "CÓD. OS" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "CÓD. OS = " & txtCodBarra.Text
            REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
        Else
            REL_Cons_Venda_Prod.rfCons2.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
        End If

    ElseIf cboCriterioPrinc.Text = "SERVIÇOS" Then
        If txtCodProduto.Text <> "" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "SERVIço = " & cboDescricao.Text
        End If

    ElseIf cboCriterioPrinc.Text = "SERVIÇOS/MENSAL" Then
        If txtCodProduto.Text <> "" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "SERVIço = " & cboDescricao.Text
            REL_Cons_Venda_Prod.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
        Else
            REL_Cons_Venda_Prod.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
        End If

    ElseIf cboCriterioPrinc.Text = "SERVIÇOS/PERÍODO" Then
        If txtCodProduto.Text <> "" Then
            REL_Cons_Venda_Prod.rfCons2.Caption = "SERVIço = " & cboDescricao.Text
            REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
        Else
            REL_Cons_Venda_Prod.rfCons2.Caption = "Inicio/Final = " & mskInicio.Text & " até " & mskFim.Text
        End If

    End If

End If

REL_Cons_Venda_Prod.dfQuant.Caption = lblQtda.Caption
REL_Cons_Venda_Prod.dfTotal.Caption = Format(lblTotal.Caption, "##,##0.00")

'REL_Cons_Venda_Prod.Relatorio.NomeImpressora = var_Impressora
REL_Cons_Venda_Prod.Relatorio.Ativar
Unload REL_Cons_Venda_Prod

Me.Show 1
End Sub

Public Sub cmdLocalizar_Click()
Dim INDICE As String

totalRegistros = "0"

'INDICE
If cboTipo.Text = "POR PRODUTOS" Then
   If cboIndice.Text = "QUANT." Then
      INDICE = "quantidade ;"
   ElseIf cboIndice.Text = "PRODUTO" Then
      INDICE = "produtos.descricao ;"
   ElseIf cboIndice.Text = "DATA" Then
      INDICE = "pedidos_itens.data ;"
   ElseIf cboIndice.Text = "PEDIDO" Then
      INDICE = "pedidos_itens.cod_pedido ;"
   Else
      INDICE = "produtos.descricao ;"
   End If
End If
If cboTipo.Text = "POR SERVIÇOS" Then
   If cboIndice.Text = "QUANT." Then
      INDICE = "s.quantidade ;"
   ElseIf cboIndice.Text = "PRODUTO" Then
      INDICE = "s.descricao ;"
   ElseIf cboIndice.Text = "DATA" Then
      INDICE = "s.data ;"
   ElseIf cboIndice.Text = "PEDIDO" Then
      INDICE = "s.cod_os ;"
   Else
      INDICE = "s.descricao ;"
   End If
End If

sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.data as varData, pedidos_itens.cod_pedido as varCodPed, pedidos_itens.cod_produto as varCodProd, produtos.descricao as varNome, produtos.fabricante as varFab, produtos.tamanho as varTam, produtos.REF as varRef, pedidos_itens.preco as varValor, pedidos_itens.quantidade as varQuant, pedidos_itens.SUBTOTAL as varSubtotal, pedidos_itens.Desconto as varDesc, pedidos_itens.Total as varTotal, ISNULL(OS.COD_OS, 0) AS var_CodOS, produtos.COD_BARRA as varCodBarra " & _
        "FROM pedidos_itens INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo LEFT OUTER JOIN OS ON pedidos.COD_PEDIDO = OS.COD_PEDIDO " & _
        "WHERE pedidos_itens.cancelado = 0 AND pedidos.tipo_pedido <> 'ORÇAMENTO'"
   
If cboTipo.Text = "POR PRODUTOS" Then

            'TODOS
             If cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "TODOS" Then
                If txtCodProduto.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.codigo = " & txtCodProduto.Text & " " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "REFERÊNCIA" And cboCriterioPrinc.Text = "TODOS" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.REF = '" & cboDescricao.Text & "' " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "FABRICANTE" And cboCriterioPrinc.Text = "TODOS" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.FABRICANTE = '" & cboDescricao.Text & "' " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "TODOS" Then
                If cboDescricao.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.CATEGORIA = '" & cboDescricao.Text & "' " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "TODOS" Then
                If txtCodBarra.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.cod_barra = '" & txtCodBarra.Text & "' " & _
                       "ORDER BY " & INDICE
                       
            'MENSAL
             ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "MENSAL" Then
                If cboDescricao.Text = "" Then Exit Sub
                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.descricao = '" & cboDescricao.Text & "' and (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") " & _
                       "ORDER BY " & INDICE
                
             ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "MENSAL" Then
                If txtCodBarra.Text = "" Then Exit Sub
                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.cod_barra = '" & txtCodBarra.Text & "' and (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "REFERÊNCIA" And cboCriterioPrinc.Text = "MENSAL" Then
                If cboDescricao.Text = "" Then Exit Sub
                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.REF = '" & cboDescricao.Text & "' and (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "FABRICANTE" And cboCriterioPrinc.Text = "MENSAL" Then
                If cboDescricao.Text = "" Then Exit Sub
                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.FABRICANTE = '" & cboDescricao.Text & "' and (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "MENSAL" Then
                If cboDescricao.Text = "" Then Exit Sub
                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.CATEGORIA = '" & cboDescricao.Text & "' and (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") " & _
                       "ORDER BY " & INDICE
            'PERÍODO
             ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "PERÍODO" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.descricao = '" & cboDescricao.Text & "' and (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                
             ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "PERÍODO" Then
                If txtCodBarra.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.cod_barra = '" & txtCodBarra.Text & "' and (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "REFERÊNCIA" And cboCriterioPrinc.Text = "PERÍODO" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.REF = '" & cboDescricao.Text & "' and (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "FABRICANTE" And cboCriterioPrinc.Text = "PERÍODO" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.FABRICANTE = '" & cboDescricao.Text & "' and (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "PERÍODO" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.CATEGORIA = '" & cboDescricao.Text & "' and (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE

            'DATA
             ElseIf cboCriterioSec.Text = "DESCRIÇÃO" And cboCriterioPrinc.Text = "DATA" Then
                If txtCodProduto.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.CODIGO = " & txtCodProduto.Text & " and (pedidos_itens.data = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                
             ElseIf cboCriterioSec.Text = "CÓD. BARRA" And cboCriterioPrinc.Text = "DATA" Then
                If txtCodBarra.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.cod_barra = '" & txtCodBarra.Text & "' and (pedidos_itens.data = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "REFERÊNCIA" And cboCriterioPrinc.Text = "DATA" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.REF = '" & cboDescricao.Text & "' and (pedidos_itens.data = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "FABRICANTE" And cboCriterioPrinc.Text = "DATA" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.FABRICANTE = '" & cboDescricao.Text & "' and (pedidos_itens.data = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
                       
             ElseIf cboCriterioSec.Text = "CATEGORIA" And cboCriterioPrinc.Text = "DATA" Then
                If cboDescricao.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.CATEGORIA = '" & cboDescricao.Text & "' and (pedidos_itens.data = CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
            'PRODUTO/MENSAL
             ElseIf cboCriterioPrinc.Text = "PRODUTO/MENSAL" Then
                If txtCodProduto.Text = "" Then Exit Sub
                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
                sSQL = sSQL & " and produtos.codigo = " & txtCodProduto.Text & " and (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") " & _
                       "ORDER BY " & INDICE
            'PRODUTO/PERÍODO
             ElseIf cboCriterioPrinc.Text = "PRODUTO/PERÍODO" Then
                If txtCodProduto.Text = "" Then Exit Sub
                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
                sSQL = sSQL & " and produtos.codigo = " & txtCodProduto.Text & " and (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) " & _
                       "ORDER BY " & INDICE
            End If
            
    'Debug.Print sSQL

        
ElseIf cboTipo.Text = "POR SERVIÇOS" Then
   Dim sBase As String
   sBase = "SELECT s.codigo, OS.COD_OS AS varCodPed, OS.DATA_TERMINO AS varData, s.descricao AS varNome, " & _
           "s.preco AS varValor, s.quantidade AS varQuant, s.subtotal AS varSubtotal, " & _
           "s.desconto AS varDesc, s.total AS varTotal, ISNULL(OS.COD_PEDIDO, 0) AS var_CodOS, s.cod_servico AS varCodServ " & _
           "FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS "

   If cboCriterioPrinc.Text = "TODOS" Then
      If cboCriterioSec.Text = "DESCRIÇÃO" Then
         If cboDescricao.Text = "" Or txtCodProduto.Text = "" Then Exit Sub
         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " ORDER BY " & INDICE
      ElseIf cboCriterioSec.Text = "CÓD. OS" Then
         If txtCodBarra.Text = "" Then Exit Sub
         sSQL = sBase & "WHERE OS.COD_OS = " & Val(txtCodBarra.Text) & " ORDER BY " & INDICE
      Else
         sSQL = sBase & "ORDER BY " & INDICE
      End If
   ElseIf cboCriterioPrinc.Text = "MENSAL" Then
      If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
      If cboCriterioSec.Text = "DESCRIÇÃO" Then
         If cboDescricao.Text = "" Or txtCodProduto.Text = "" Then Exit Sub
         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " AND MONTH(OS.DATA_TERMINO) = " & cboMes.ListIndex + 1 & " AND YEAR(OS.DATA_TERMINO) = " & cboAno & " ORDER BY " & INDICE
      ElseIf cboCriterioSec.Text = "CÓD. OS" Then
         If txtCodBarra.Text = "" Then Exit Sub
         sSQL = sBase & "WHERE OS.COD_OS = " & Val(txtCodBarra.Text) & " AND MONTH(OS.DATA_TERMINO) = " & cboMes.ListIndex + 1 & " AND YEAR(OS.DATA_TERMINO) = " & cboAno & " ORDER BY " & INDICE
      Else
         sSQL = sBase & "WHERE MONTH(OS.DATA_TERMINO) = " & cboMes.ListIndex + 1 & " AND YEAR(OS.DATA_TERMINO) = " & cboAno & " ORDER BY " & INDICE
      End If
   ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
      If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
      If cboCriterioSec.Text = "DESCRIÇÃO" Then
         If cboDescricao.Text = "" Or txtCodProduto.Text = "" Then Exit Sub
         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " AND (OS.DATA_TERMINO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (OS.DATA_TERMINO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
      ElseIf cboCriterioSec.Text = "CÓD. OS" Then
         If txtCodBarra.Text = "" Then Exit Sub
         sSQL = sBase & "WHERE OS.COD_OS = " & Val(txtCodBarra.Text) & " AND (OS.DATA_TERMINO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (OS.DATA_TERMINO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
      Else
         sSQL = sBase & "WHERE (OS.DATA_TERMINO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (OS.DATA_TERMINO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
      End If
   ElseIf cboCriterioPrinc.Text = "SERVIÇOS/MENSAL" Then
      If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
      If txtCodProduto.Text <> "" Then
         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " AND MONTH(OS.DATA_TERMINO) = " & cboMes.ListIndex + 1 & " AND YEAR(OS.DATA_TERMINO) = " & cboAno & " ORDER BY " & INDICE
      Else
         sSQL = sBase & "WHERE MONTH(OS.DATA_TERMINO) = " & cboMes.ListIndex + 1 & " AND YEAR(OS.DATA_TERMINO) = " & cboAno & " ORDER BY " & INDICE
      End If
   ElseIf cboCriterioPrinc.Text = "SERVIÇOS/PERÍODO" Then
      If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
      If txtCodProduto.Text <> "" Then
         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " AND (OS.DATA_TERMINO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (OS.DATA_TERMINO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
      Else
         sSQL = sBase & "WHERE (OS.DATA_TERMINO >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (OS.DATA_TERMINO <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) ORDER BY " & INDICE
      End If
   ElseIf cboCriterioPrinc.Text = "SERVIÇOS" Then
      If txtCodProduto.Text <> "" Then
         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " ORDER BY " & INDICE
      Else
         sSQL = sBase & "ORDER BY " & INDICE
      End If
   End If
End If
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

Debug.Print sSQL

If cboTipo.Text = "POR SERVIÇOS" Then
   FormatarGrid_Servicos r
Else
   FormatarGrid_ProdDetalhado r
End If
'FormatarGrid_Produtos r
printSQL = sSQL

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
   
'FORMATAR O GRID
With Grid
   .Clear
   .Cols = 7
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   .ColWidth(3) = 0
   .ColWidth(4) = 0
   .ColWidth(5) = 0
   .ColWidth(6) = 0
End With

PreencherCriterio
cboCriterioPrinc.ListIndex = 0

PreencherTipoConsulta
cboTipo.ListIndex = 0

PreencherCriterioSec
cboCriterioSec.ListIndex = 1

PreencherIndice
cboIndice.ListIndex = 2

Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   posX = x
   Label3 = posX
   If Label3.Caption > 0 And Label3.Caption < 149 Then Grid.ToolTipText = ""
   If Label3.Caption > 150 And Label3.Caption < 930 Then Grid.ToolTipText = "Dê um duplo-clique para exibir os itens do Pedido."
   If Label3.Caption > 931 And Label3.Caption < 7230 Then Grid.ToolTipText = ""
   If Label3.Caption > 7231 And Label3.Caption < 8355 Then Grid.ToolTipText = "Dê um duplo-clique para exibir a forma de pgto."
   If Label3.Caption > 8356 And Label3.Caption < 9555 Then Grid.ToolTipText = ""
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
      cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFim.SetFocus
      SelectControl mskFim
   End If
End If
End Sub

Private Sub mskInicio_GotFocus()
   SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub

Sub FormatarGrid_Vendas(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1220
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "TIPO"
      .TextMatrix(0, 7) = "TIPO"
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.rows - 1, 4) = Format(rTabela("var_total"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = rTabela("tipo_pagamento")
            .TextMatrix(.rows - 1, 6) = rTabela("pagamento")
            .TextMatrix(.rows - 1, 7) = rTabela("tipo_pedido")
            
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)

picAguarde.Visible = False
End Sub

Sub FormatarGrid_VendasComEntrada(rTabela As ADODB.Recordset)
   Dim i As Integer
picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 3600
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 800
      .ColWidth(7) = 1100
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "ENTRADA"
      .TextMatrix(0, 5) = "VALOR"
      .TextMatrix(0, 6) = "FORMA"
      .TextMatrix(0, 7) = "TIPO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      .Redraw = False
      
      i = 1
      
            '.TextMatrix(.Rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            '.TextMatrix(.Rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            '.TextMatrix(.Rows - 1, 3) = UCase(rTabela("nome"))
            '.TextMatrix(.Rows - 1, 4) = Format(rTabela("var_total"), ocMONEY)
            '.TextMatrix(.Rows - 1, 5) = rTabela("tipo_pagamento")
            '.TextMatrix(.Rows - 1, 6) = rTabela("pagamento")
            '.TextMatrix(.Rows - 1, 7) = rTabela("tipo_pedido")

      
      
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("var_codped"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.rows - 1, 4) = Format(rTabela("valor_final"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = Format(rTabela("var_total"), ocMONEY)
            .TextMatrix(.rows - 1, 6) = rTabela("tipo_pagamento")
            .TextMatrix(.rows - 1, 7) = rTabela("pagamento")
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   lblTotal.Caption = Format(SomaGrid(Grid, 5), "##,##0.00")
picAguarde.Visible = False
End Sub

Private Sub mskInicio_LostFocus()
   If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
      mskInicio.Mask = ""
      mskInicio.Text = ""
      Exit Sub
   Else
      If IsDate(mskInicio.Text) Then
         If mskFim.Visible = True Then mskFim.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskInicio.SetFocus
         SelectControl mskInicio
      End If
   End If
End Sub

