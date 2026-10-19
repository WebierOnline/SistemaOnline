VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Saida_Consulta_PorProdutosAgrupadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE SAÍDAS"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "Saida_Consulta_PorProdutosAgrupadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   60
      ScaleHeight     =   2565
      ScaleWidth      =   10905
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
      Top             =   840
      Width           =   10935
      Begin VB.Frame Frame8 
         Caption         =   "Consulta"
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
         Height          =   2415
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   10695
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
            Height          =   2055
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox cboCriterioSec 
               Height          =   315
               Left            =   120
               TabIndex        =   30
               Top             =   420
               Width           =   2175
            End
            Begin VB.ComboBox cboCriterioPrinc 
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Top             =   1020
               Width           =   2175
            End
            Begin VB.ComboBox cboIndice 
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   1620
               Width           =   2175
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Consultar por:"
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Filtrar por:"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   780
               Width           =   750
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Organizar por:"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   1390
               Width           =   990
            End
         End
         Begin VB.Frame frmFiltro1 
            Height          =   1500
            Left            =   2640
            TabIndex        =   9
            Top             =   240
            Width           =   7935
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   1500
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   1020
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMes 
               Height          =   315
               Left            =   120
               TabIndex        =   12
               Top             =   1020
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboDescricao 
               Height          =   315
               Left            =   120
               TabIndex        =   11
               Top             =   420
               Visible         =   0   'False
               Width           =   5955
            End
            Begin VB.TextBox txtCodBarra 
               Height          =   315
               Left            =   120
               TabIndex        =   10
               Top             =   420
               Visible         =   0   'False
               Width           =   2355
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario1 
               Height          =   315
               Left            =   1080
               TabIndex        =   24
               Tag             =   "Calendario"
               Top             =   1020
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
               MICON           =   "Saida_Consulta_PorProdutosAgrupadas.frx":23D2
               PICN            =   "Saida_Consulta_PorProdutosAgrupadas.frx":23EE
               PICH            =   "Saida_Consulta_PorProdutosAgrupadas.frx":4741
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
               Left            =   2460
               TabIndex        =   25
               Tag             =   "Calendario"
               Top             =   1020
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
               MICON           =   "Saida_Consulta_PorProdutosAgrupadas.frx":6A94
               PICN            =   "Saida_Consulta_PorProdutosAgrupadas.frx":6AB0
               PICH            =   "Saida_Consulta_PorProdutosAgrupadas.frx":8E03
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   1020
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
               Left            =   1500
               TabIndex        =   27
               Top             =   1020
               Visible         =   0   'False
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin VB.ComboBox cboFormaSaida 
               Height          =   315
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   420
               Width           =   1700
            End
            Begin VB.Label lblAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1500
               TabIndex        =   16
               Top             =   780
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   780
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblDescricao 
               AutoSize        =   -1  'True
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   180
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label lblFormaSaida 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma:"
               Height          =   195
               Left            =   2520
               TabIndex        =   28
               Top             =   180
               Width           =   480
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdSair 
            Height          =   495
            Left            =   9240
            TabIndex        =   32
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
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
            MICON           =   "Saida_Consulta_PorProdutosAgrupadas.frx":B156
            PICN            =   "Saida_Consulta_PorProdutosAgrupadas.frx":B172
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
            Height          =   495
            Left            =   7800
            TabIndex        =   33
            Top             =   1800
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Imprimir"
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
            MICON           =   "Saida_Consulta_PorProdutosAgrupadas.frx":B48C
            PICN            =   "Saida_Consulta_PorProdutosAgrupadas.frx":B4A8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   495
            Left            =   6300
            TabIndex        =   34
            Top             =   1800
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
            MICON           =   "Saida_Consulta_PorProdutosAgrupadas.frx":B7C2
            PICN            =   "Saida_Consulta_PorProdutosAgrupadas.frx":B7DE
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
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      Picture         =   "Saida_Consulta_PorProdutosAgrupadas.frx":C0B8
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   10905
      TabIndex        =   0
      Top             =   60
      Width           =   10935
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SAÍDA DE PRODUTOS AGRUPADAS"
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
         Left            =   1380
         TabIndex        =   1
         Top             =   120
         Width           =   5970
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   240
         Picture         =   "Saida_Consulta_PorProdutosAgrupadas.frx":D0F0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   9810
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 9427-5280"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "14:49"
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
      Height          =   5955
      Left            =   60
      TabIndex        =   17
      Top             =   3480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10504
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirEntradas 
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   9480
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "MOSTRAR SAÍDAS"
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
      MICON           =   "Saida_Consulta_PorProdutosAgrupadas.frx":13936
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
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
      Left            =   8520
      TabIndex        =   5
      Top             =   9540
      Visible         =   0   'False
      Width           =   780
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
      Left            =   9360
      TabIndex        =   4
      Top             =   9480
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   5700
      TabIndex        =   2
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Saida_Consulta_PorProdutosAgrupadas"
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
Private Sub FormatarGrid_Produtos(rTabela As ADODB.Recordset)
   Dim i As Integer

picAguarde.Visible = True
DoEvents
   With Grid
      .Clear
      .Cols = 5
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 900
      .ColWidth(2) = 1300
      .ColWidth(3) = 5395
      .ColWidth(4) = 850
      
      .TextMatrix(0, 1) = "CÓD."
      .TextMatrix(0, 2) = "CÓD. BARRA"
      .TextMatrix(0, 3) = "DESCRIÇÃO"
      .TextMatrix(0, 4) = "QTDE"
      
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
            .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("COD_PRODUTO"))
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("cod_barra"))
            .TextMatrix(.rows - 1, 3) = rTabela("DESCRICAO") & " /  " & rTabela("FABRICANTE")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("VsOMAQUANT"))
            
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      .Redraw = True
      .Row = 0
      .Col = 0
   End With
   
   lblQtda.Caption = Format(SomaGrid(Grid, 4), ocPESO)
picAguarde.Visible = False
End Sub
Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
cboDescricao.Text = ""
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

Private Sub PreencherConsulta()
cboCriterioSec.AddItem "DESCRIÇÃO"
cboCriterioSec.AddItem "CÓD. BARRA"
cboCriterioSec.AddItem "CÓD. PRODUTO"
cboCriterioSec.AddItem "CATEGORIA"
End Sub

Private Sub AtualizarCriterioSec()
If cboCriterioSec.Text = "" Then cboCriterioSec.Text = "DESCRIÇÃO"

lblDescricao.Visible = False
txtCodBarra.Visible = False
cboDescricao.Visible = False

If cboCriterioSec.Text = "CÓD. BARRA" Then
   lblDescricao.Caption = "Cod. Barra:"
   lblDescricao.Visible = True
   txtCodBarra.Visible = True
   txtCodBarra.Text = ""
   cboFormaSaida.Left = 2520
   lblFormaSaida.Left = 2520
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" Then
   lblDescricao.Caption = "Cód. Produto:"
   lblDescricao.Visible = True
   txtCodBarra.Visible = True
   txtCodBarra.Text = ""
   cboFormaSaida.Left = 2520
   lblFormaSaida.Left = 2520
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   lblDescricao.Caption = "Descrição:"
   lblDescricao.Visible = True
   cboDescricao.Visible = True
   cboFormaSaida.Left = 6120
   lblFormaSaida.Left = 6120
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   lblDescricao.Caption = "Categoria:"
   lblDescricao.Visible = True
   cboDescricao.Visible = True
   cboFormaSaida.Left = 6120
   lblFormaSaida.Left = 6120
End If
End Sub

Private Sub cboCriterioSec_Click()
'Não espera perder o foco - exibe os objetos e dá foco na hora que escolhe.
'SetFocus fica só aqui (não em LostFocus) pra evitar reentrância (erro 5)
AtualizarCriterioSec
DoEvents
If txtCodBarra.Visible Then
   txtCodBarra.SetFocus
ElseIf cboDescricao.Visible Then
   cboDescricao.SetFocus
End If
End Sub

Private Sub cboCriterioSec_GotFocus()
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
   cmdCalendario1.Visible = False
   cmdCalendario2.Visible = False
   mskInicio.Visible = False
   mskFim.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
   lblMes.Visible = True
   lblAno.Visible = True
   lblMes.Caption = "Mês"
   lblAno.Caption = "Ano"
   cmdCalendario1.Visible = False
   cmdCalendario2.Visible = False
   mskInicio.Visible = False
   mskFim.Visible = False
   cboMes.Visible = True
   cboAno.Visible = True
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
   lblMes.Visible = True
   lblAno.Visible = True
   lblMes.Caption = "Inicio"
   lblAno.Caption = "Final"
   cmdCalendario1.Visible = True
   cmdCalendario2.Visible = True
   mskInicio.Visible = True
   mskFim.Visible = True
   cboMes.Visible = False
   cboAno.Visible = False
Else
End If
End Sub

Private Sub cboCriterioPrinc_Click()
'Não espera perder o foco - exibe os objetos na hora que escolhe. SetFocus
'fica só aqui (não em LostFocus) pra evitar reentrância: dar SetFocus tira
'o foco do combo, o que dispararia o LostFocus de verdade de novo, numa
'hora ruim (erro 5)
AtualizarCriterioPrinc
DoEvents
If cboMes.Visible Then
   cboMes.SetFocus
ElseIf mskInicio.Visible Then
   mskInicio.SetFocus
End If
End Sub

Private Sub cboCriterioPrinc_GotFocus()
cboCriterioPrinc.Clear
cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "PERÍODO"
moCombo.AttachTo cboCriterioPrinc
End Sub


Private Sub cboCriterioPrinc_LostFocus()
AtualizarCriterioPrinc
End Sub

Private Sub cboCriterioPrinc_Validate(Cancel As Boolean)
If cboCriterioPrinc.Text = "ESPECIFICO/MENSAL" Then
   lblMes.Visible = True
   cboMes.Visible = True
   lblAno.Visible = True
   cboAno.Visible = True
   cboDescricao.Visible = True
   lblDescricao.Visible = True
End If
End Sub

Private Sub cboDescricao_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

cboDescricao.Clear

If cboCriterioSec.Text = "DESCRIÇÃO" Then
   sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)

   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      r.MoveNext
   Loop
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"
   Set r = dbData.OpenRecordset(sSQL)

   Do While Not r.EOF
      cboDescricao.AddItem ValidateNull(r("categoria"))
      r.MoveNext
   Loop
Else
   Exit Sub
End If

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboDescricao
End Sub

Private Sub cboIndice_GotFocus()
cboIndice.Clear
cboIndice.AddItem "QUANT."
cboIndice.AddItem "DESCRIÇÃO"
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

Private Sub chameleonButton1_Click()
Dim INDICE As String
Dim sPeriodoFiltro As String
Dim sSecFiltro As String
Dim sFormaProdQuant As String
Dim sAjusteSaidaBranch As String
Dim sVendaBranch As String
Dim sUnion As String

If cboCriterioPrinc.Text = "" Then Exit Sub
If cboCriterioSec.Text = "" Then Exit Sub

If cboIndice.Text = "QUANT." Then
   INDICE = "SUM(mv.QUANT) DESC;"
ElseIf cboIndice.Text = "DESCRIÇÃO" Then
   INDICE = "mv.DESCRICAO;"
Else
   INDICE = "mv.DESCRICAO;"
End If

'Filtro de período (TODOS/MENSAL/PERÍODO) - aplicado na view unificada (mv.data)
sPeriodoFiltro = ""
If cboCriterioPrinc.Text = "MENSAL" Then
   If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
   sPeriodoFiltro = "AND (MONTH(mv.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(mv.data) = " & cboAno & ") "
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
   If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
   sPeriodoFiltro = "AND (mv.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (mv.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
End If

'Filtro por produto (DESCRIÇÃO/CÓD.BARRA/CÓD.PRODUTO/CATEGORIA) - na view unificada
sSecFiltro = ""
If cboCriterioSec.Text = "CÓD. BARRA" Then
   If txtCodBarra.Text = "" Then Exit Sub
   sSecFiltro = "AND (mv.cod_barra = '" & txtCodBarra.Text & "') "
ElseIf cboCriterioSec.Text = "CÓD. PRODUTO" Then
   If txtCodBarra.Text = "" Or Not IsNumeric(txtCodBarra.Text) Then Exit Sub
   sSecFiltro = "AND (mv.COD_PRODUTO = " & txtCodBarra.Text & ") "
ElseIf cboCriterioSec.Text = "DESCRIÇÃO" Then
   If cboDescricao.Text = "" Then Exit Sub
   sSecFiltro = "AND (mv.DESCRICAO = '" & cboDescricao.Text & "') "
ElseIf cboCriterioSec.Text = "CATEGORIA" Then
   If cboDescricao.Text = "" Then Exit Sub
   sSecFiltro = "AND (mv.categoria = '" & cboDescricao.Text & "') "
End If

'HISTÓRICO DE SAÍDA = as 3 formas que subtraem quant_estoque (mesma lógica do
'Saida_Consulta_PorProdutos.frm): AJUSTE/SAÍDA vêm de Produtos_Quant, VENDA vem
'direto de pedidos/pedidos_itens (não tem log em Produtos_Quant)
sAjusteSaidaBranch = ""
sVendaBranch = ""

If cboFormaSaida.Text = "" Or cboFormaSaida.Text = "TODOS" Or cboFormaSaida.Text = "AJUSTE" Or cboFormaSaida.Text = "SAÍDA" Then
   sFormaProdQuant = ""
   If cboFormaSaida.Text = "AJUSTE" Then
      sFormaProdQuant = "AND (Produtos_Quant.FORMA = 'AJUSTE') "
   ElseIf cboFormaSaida.Text = "SAÍDA" Then
      sFormaProdQuant = "AND (Produtos_Quant.FORMA = 'SAÍDA') "
   End If

   'ABS() porque QUANT tem 2 convenções diferentes de sinal pra REMOÇÃO
   '(Produtos_AdicionarQuant.frm grava sempre positivo + TIPO, Produtos_Estoque_Simples.frm
   'grava negativo direto) - sem o ABS a soma ficaria errada misturando os dois sinais
   sAjusteSaidaBranch = "SELECT Produtos_Quant.COD_PRODUTO AS COD_PRODUTO, ABS(Produtos_Quant.QUANT) AS QUANT, produtos.descricao AS DESCRICAO, produtos.FABRICANTE AS FABRICANTE, produtos.cod_barra AS cod_barra, produtos.categoria AS categoria, Produtos_Quant.data AS data " & _
      "FROM Produtos_Quant INNER JOIN produtos ON Produtos_Quant.COD_PRODUTO = produtos.CODIGO " & _
      "WHERE (Produtos_Quant.TIPO = 'REMOÇÃO') " & sFormaProdQuant
End If

If cboFormaSaida.Text = "" Or cboFormaSaida.Text = "TODOS" Or cboFormaSaida.Text = "VENDA" Then
   sVendaBranch = "SELECT pedidos_itens.cod_produto AS COD_PRODUTO, pedidos_itens.quantidade AS QUANT, produtos.descricao AS DESCRICAO, produtos.FABRICANTE AS FABRICANTE, produtos.cod_barra AS cod_barra, produtos.categoria AS categoria, pedidos.DATA_COMPRA AS data " & _
      "FROM pedidos_itens INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      "INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
      "WHERE (pedidos.CANCELADO = 0) AND (pedidos.tipo_pedido <> 'ORÇAMENTO') AND (pedidos.tipo_pagamento IN ('À Vista', 'À prazo')) AND (ISNULL(pedidos_itens.tipo_venda,'') <> 'OFICINA') "
End If

If sAjusteSaidaBranch <> "" And sVendaBranch <> "" Then
   sUnion = sAjusteSaidaBranch & "UNION ALL " & sVendaBranch
ElseIf sAjusteSaidaBranch <> "" Then
   sUnion = sAjusteSaidaBranch
ElseIf sVendaBranch <> "" Then
   sUnion = sVendaBranch
Else
   Exit Sub
End If

sSQL = "SELECT mv.COD_PRODUTO, SUM(mv.QUANT) AS vSomaQuant, mv.DESCRICAO, mv.FABRICANTE, mv.cod_barra " & _
    "FROM (" & sUnion & ") AS mv " & _
    "WHERE (1 = 1) " & sPeriodoFiltro & sSecFiltro & _
    "GROUP BY mv.COD_PRODUTO, mv.DESCRICAO, mv.FABRICANTE, mv.cod_barra " & _
    "ORDER BY " & INDICE

Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid_Produtos r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

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


Private Sub cmdExibirEntradas_Click()
If Grid.Col = 0 Or Grid.Row = 0 Then Exit Sub
If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then Exit Sub
If Val(Grid.TextMatrix(Grid.Row, 1)) = 0 Then Exit Sub

Dim lMes As Integer
Dim sAno As String
Dim sDataIni As String
Dim sDataFim As String

lMes = 0
sAno = ""
sDataIni = ""
sDataFim = ""

If cboCriterioPrinc.Text = "MENSAL" Then
   lMes = cboMes.ListIndex + 1
   sAno = cboAno.Text
ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
   sDataIni = mskInicio.Text
   sDataFim = mskFim.Text
End If

Saida_Consulta_PorProdutos.FiltrarPorProdutoEForma Val(Grid.TextMatrix(Grid.Row, 1)), cboFormaSaida.Text, cboCriterioPrinc.Text, lMes, sAno, sDataIni, sDataFim
Me.Hide
Saida_Consulta_PorProdutos.Show
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

Private Sub cmdImprimir_Click()
ShowMsg "Impressão ainda não implementada pra essa tela.", vbInformation
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub



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

PreencherConsulta
cboCriterioSec.ListIndex = 1   'CÓD. BARRA

cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "PERÍODO"
cboCriterioPrinc.ListIndex = 0   'TODOS

cboFormaSaida.AddItem "TODOS"
cboFormaSaida.AddItem "VENDA"
cboFormaSaida.AddItem "AJUSTE"
cboFormaSaida.AddItem "SAÍDA"
cboFormaSaida.ListIndex = 0

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
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

