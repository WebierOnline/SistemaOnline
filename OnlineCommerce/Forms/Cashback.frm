VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Cashback 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CASHBACK"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   Icon            =   "Cashback.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSituacao 
      Caption         =   "Situação"
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
      Height          =   1395
      Left            =   2880
      TabIndex        =   26
      Top             =   6360
      Width           =   1575
      Begin VB.OptionButton optAtivos 
         Caption         =   "Ativos"
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
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optAbatidos 
         Caption         =   "Abatidos"
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
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optInutilizado 
         Caption         =   "Inutilizados"
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
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optSituacaoTodos 
         Caption         =   "Todos"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmIndice 
      Caption         =   "Organizar"
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
      Height          =   1395
      Left            =   1380
      TabIndex        =   23
      Top             =   6360
      Width           =   1455
      Begin VB.OptionButton optIndiceCliente 
         Caption         =   "Cliente"
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
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optIndiceValidade 
         Caption         =   "Validade"
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
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame frmCriterios 
      Caption         =   "Tipo"
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
      Height          =   1395
      Left            =   60
      TabIndex        =   18
      Top             =   6360
      Width           =   1275
      Begin VB.OptionButton optMensal 
         Caption         =   "Mensal"
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
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
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
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Período"
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
         TabIndex        =   20
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Cliente"
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Preencha"
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   4500
      TabIndex        =   4
      Top             =   6360
      Width           =   8235
      Begin VB.ComboBox cboMes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   1755
      End
      Begin VB.ComboBox cboAno 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.ComboBox cboNome 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   5715
      End
      Begin VB.TextBox txtCodClie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4980
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin ChamaleonBtn.chameleonButton cmdCal2 
         Height          =   315
         Left            =   5520
         TabIndex        =   9
         Tag             =   "Calendario"
         Top             =   1020
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cashback.frx":23D2
         PICN            =   "Cashback.frx":23EE
         PICH            =   "Cashback.frx":4741
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Tag             =   "Calendario"
         Top             =   1020
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cashback.frx":6A94
         PICN            =   "Cashback.frx":6AB0
         PICH            =   "Cashback.frx":8E03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin MSMask.MaskEdBox mskFim 
         Height          =   315
         Left            =   4620
         TabIndex        =   11
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   3300
         TabIndex        =   12
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFim 
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
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
         Height          =   255
         Left            =   4620
         TabIndex        =   17
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lblInicio 
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
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
         Height          =   180
         Left            =   3360
         TabIndex        =   16
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblMes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
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
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   800
         Width           =   360
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
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
         Height          =   195
         Left            =   1920
         TabIndex        =   14
         Top             =   800
         Width           =   345
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cliente:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   1470
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   14325
      TabIndex        =   0
      Top             =   60
      Width           =   14355
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Cashback.frx":B156
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cashback"
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
         TabIndex        =   1
         Top             =   180
         Width           =   1530
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   7815
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21246
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:43"
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
   Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
      Height          =   5475
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   9657
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibir 
      Height          =   375
      Left            =   12840
      TabIndex        =   31
      Top             =   6420
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
      MICON           =   "Cashback.frx":114C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEstaSemana 
      Height          =   375
      Left            =   12840
      TabIndex        =   32
      Top             =   6840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Lembrente"
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
      MICON           =   "Cashback.frx":114E2
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
      Height          =   375
      Left            =   12840
      TabIndex        =   33
      Top             =   7260
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "Cashback.frx":114FE
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
Attribute VB_Name = "Produtos_Cashback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Private printSQL As String

Private Sub FormatarGridConsulta(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Consulta
      .Clear
      .Cols = 12
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 600
      .ColWidth(2) = 700
      .ColWidth(3) = 4000
      .ColWidth(4) = 1250
      .ColWidth(5) = 1050
      .ColWidth(6) = 1100
      .ColWidth(7) = 1100
      .ColWidth(8) = 900
      .ColWidth(9) = 900
      .ColWidth(10) = 1000
      .ColWidth(11) = 900

      
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "VENDA"
      .TextMatrix(0, 3) = "CLIENTE"
      .TextMatrix(0, 4) = "CELULAR"
      .TextMatrix(0, 5) = "VALIDADE"
      .TextMatrix(0, 6) = "CASHBACK"
      .TextMatrix(0, 7) = "ABATIDO"
      .TextMatrix(0, 8) = "DATA"
      .TextMatrix(0, 9) = "CÓD.ABATIDO"
      .TextMatrix(0, 10) = "INVÁLIDO"
      .TextMatrix(0, 11) = "FUNC."

      .Redraw = False
      
      i = 1
      
      .ColAlignment(1) = 2
      .ColAlignment(2) = 2
      '.ColAlignment(3) = 1

'Pedidos_Cashback.CODIGO AS varCod, Pedidos_Cashback.COD_PEDIDO AS varCodPed, Pedidos_Cashback.COD_CLIENTE AS varCodCli, cliente.Nome AS varCli,
'Pedidos_Cashback.VALOR_VENDA AS varVenda, Pedidos_Cashback.VALOR_CASHBACK AS varVlrCash, Pedidos_Cashback.VALOR_ABATIDO AS varVlrAbatido, Pedidos_Cashback.ABATIDO AS varAbatido, Pedidos_Cashback.DATA_ABATIDO AS varDataAbatido,
'Pedidos_Cashback.COD_PEDIDOABATIDO AS varCodPedAbatido, Pedidos_Cashback.VALIDADE AS varValidade, Pedidos_Cashback.INVALIDO AS varInvalido, Pedidos_Cashback.COD_FUNCIONARIO AS varCodFunc

      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("varCod")
            .TextMatrix(.rows - 1, 2) = rTabela("varCodPed") & ""
            .TextMatrix(.rows - 1, 3) = rTabela("varCli")
            .TextMatrix(.rows - 1, 4) = rTabela("varCelular")
            .TextMatrix(.rows - 1, 5) = Format(ValidateNull(rTabela("varValidade")), "dd/mm/yy")
            .TextMatrix(.rows - 1, 6) = FormatNumber(rTabela("varVlrCash"), 2)
            .TextMatrix(.rows - 1, 7) = rTabela("varAbatido")
            If IsNull(rTabela("varDataAbatido")) Then
                .TextMatrix(.rows - 1, 8) = ""
            Else
                .TextMatrix(.rows - 1, 8) = Format(rTabela("varDataAbatido"), "dd/mm/yy")
            End If
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("varCodPedAbatido"))
            .TextMatrix(.rows - 1, 10) = rTabela("varInvalido")
            .TextMatrix(.rows - 1, 11) = rTabela("varCodFunc")
            
            rTabela.MoveNext
            .rows = .rows + 1
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
      
      .rows = .rows - 1
      .Redraw = True
   End With
End Sub

Private Sub cboMes_GotFocus()
Dim vMes As Integer

cboMes.Clear
For vMes = 1 To 12
   cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
Next

moCombo.AttachTo cboMes
End Sub

Private Sub cboMes_LostFocus()
   cboAno.SetFocus
End Sub

Private Sub cboNome_GotFocus()
cboNome.Clear

sSQL = "SELECT * FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboNome.AddItem r("nome")
   cboNome.ItemData(cboNome.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing
   
moCombo.AttachTo cboNome
End Sub

Private Sub cboNome_LostFocus()
If cboNome.Text = "" Then txtCodClie.Text = "": Exit Sub
If cboNome.ListIndex = -1 Then txtCodClie.Text = "": Exit Sub
txtCodClie = cboNome.ItemData(cboNome.ListIndex)
Exit Sub
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
End Sub

Private Sub cmdEstaSemana_Click()
Dim vDiaSemanaAtual As Integer
Dim vDataDomingo As Date
Dim vDataSabado As Date

vDiaSemanaAtual = Weekday(Now)

If vDiaSemanaAtual = 1 Then
    vDataDomingo = DateAdd("d", 0, Date)
    vDataSabado = DateAdd("d", 6, Date)
ElseIf vDiaSemanaAtual = 2 Then
    vDataDomingo = DateAdd("d", -1, Date)
    vDataSabado = DateAdd("d", 5, Date)
ElseIf vDiaSemanaAtual = 3 Then
    vDataDomingo = DateAdd("d", -2, Date)
    vDataSabado = DateAdd("d", 6, Date)
ElseIf vDiaSemanaAtual = 4 Then
    vDataDomingo = DateAdd("d", -3, Date)
    vDataSabado = DateAdd("d", 3, Date)
ElseIf vDiaSemanaAtual = 5 Then
    vDataDomingo = DateAdd("d", -4, Date)
    vDataSabado = DateAdd("d", 2, Date)
ElseIf vDiaSemanaAtual = 6 Then
    vDataDomingo = DateAdd("d", -5, Date)
    vDataSabado = DateAdd("d", 1, Date)
ElseIf vDiaSemanaAtual = 7 Then
    vDataDomingo = DateAdd("d", -6, Date)
    vDataSabado = DateAdd("d", 0, Date)
End If

sSQL = "SELECT Pedidos_Cashback.CODIGO AS varCod, Pedidos_Cashback.COD_PEDIDO AS varCodPed, Pedidos_Cashback.COD_CLIENTE AS varCodCli, cliente.Nome AS varCli, cliente.Celular AS varCelular, " & _
        "Pedidos_Cashback.VALOR_VENDA AS varVenda, Pedidos_Cashback.VALOR_CASHBACK AS varVlrCash, Pedidos_Cashback.VALOR_ABATIDO AS varVlrAbatido, " & _
        "(CASE WHEN Pedidos_Cashback.ABATIDO = 1 THEN 'SIM' ELSE '' END) AS varAbatido, Pedidos_Cashback.DATA_ABATIDO AS varDataAbatido, " & _
        "Pedidos_Cashback.COD_PEDIDOABATIDO AS varCodPedAbatido, Pedidos_Cashback.VALIDADE AS varValidade, " & _
        "(CASE WHEN Pedidos_Cashback.INVALIDO = 1 THEN 'SIM' ELSE '' END) AS varInvalido, Pedidos_Cashback.COD_FUNCIONARIO AS varCodFunc " & _
"FROM Pedidos_Cashback INNER JOIN cliente ON Pedidos_Cashback.COD_CLIENTE = cliente.CODIGO " & _
"WHERE (Pedidos_Cashback.VALIDADE >= CONVERT(DATETIME, '" & Format(vDataDomingo, ocDATA) & "', 103)) AND (Pedidos_Cashback.VALIDADE <= CONVERT(DATETIME, '" & Format(vDataSabado, ocDATA) & "', 103)) AND (Pedidos_Cashback.ABATIDO = 0) AND (Pedidos_Cashback.INVALIDO = 0) " & _
"ORDER BY varValidade"
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)
printSQL = sSQL
FormatarGridConsulta r

If r.State <> 0 Then r.Close
Set r = Nothing

End Sub


Private Sub cmdImprimir_Click()
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

   Set REL_Cashback_Semana.Relatorio.Recordset = r
   
   'REL_Cashback_Semana.dfQuant.Caption = lblQtda.Caption
   'REL_Cashback_Semana.dfSubtotal.Caption = lblTotal.Caption
   
   REL_Cashback_Semana.lblTitulo.Caption = "RELATÓRIO DE CASHBACK"

   'If cboFormaPgto.Text = "TODOS" Then
   '   REL_Cashback_Semana.rfForma.Caption = "TODAS"
   'ElseIf cboFormaPgto.Text = "À VISTA" Then
   '   REL_Cashback_Semana.rfForma.Caption = "À VISTA"
   'ElseIf cboFormaPgto.Text = "À PRAZO" Then
   '   REL_Cashback_Semana.rfForma.Caption = "À PRAZO"
   'Else
   '   REL_Cashback_Semana.rfForma.Caption = "TODAS"
   'End If

   'If cboCriterioPrinc.Text = "VENDEDOR" Then
   '   REL_Cashback_Semana.rfCons1.Caption = "Vendedor = " & cboVendedor.Text & ""
   'ElseIf cboCriterioPrinc.Text = "CLIENTE" Then
   '   REL_Cashback_Semana.rfCons1.Caption = "Cliente = " & cboCliente.Text & ""
   'ElseIf cboCriterioPrinc.Text = "PERIODO" Then
   '   REL_Cashback_Semana.rfCons1.Caption = "Intervalo de " & mskInicio.Text & " à " & mskFim.Text
   'ElseIf cboCriterioPrinc.Text = "CÓDIGO" Then
   '   REL_Cashback_Semana.rfCons1.Caption = "Código = " & txtCodigo.Text & ""
   'ElseIf cboCriterioPrinc.Text = "MENSAL" Then
   '   REL_Cashback_Semana.rfCons1.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
   'ElseIf cboCriterioPrinc.Text = "TODOS" Then
   '   REL_Cashback_Semana.rfCons1.Caption = "TODOS"
   'Else
   '   REL_Cashback_Semana.rfCons1.Caption = "TODOS"
   'End If

   'If cboCriterioSec.Text = "MENSAL" Then
   '   REL_Cashback_Semana.rfCons2.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
   'End If
   
   REL_Cashback_Semana.Relatorio.Ativar
   Unload REL_Cashback_Semana

Me.Show 1
End Sub



Private Sub optCliente_Click()
lblNome.Enabled = True
cboNome.Enabled = True
lblMes.Enabled = False
cboMes.Enabled = False
lblAno.Enabled = False
cboAno.Enabled = False
lblInicio.Enabled = False
lblFim.Enabled = False
mskInicio.Enabled = False
mskFim.Enabled = False
cmdCal1.Enabled = False
cmdCal2.Enabled = False
End Sub

Private Sub optMensal_Click()
lblNome.Enabled = False
cboNome.Enabled = False
lblMes.Enabled = True
cboMes.Enabled = True
lblAno.Enabled = True
cboAno.Enabled = True
lblInicio.Enabled = False
lblFim.Enabled = False
mskInicio.Enabled = False
mskFim.Enabled = False
cmdCal1.Enabled = False
cmdCal2.Enabled = False
End Sub

Private Sub optPeriodo_Click()
lblNome.Enabled = False
cboNome.Enabled = False
lblMes.Enabled = False
cboMes.Enabled = False
lblAno.Enabled = False
cboAno.Enabled = False
lblInicio.Enabled = True
lblFim.Enabled = True
mskInicio.Enabled = True
mskFim.Enabled = True
cmdCal1.Enabled = True
cmdCal2.Enabled = True
End Sub


Private Sub optTodos_Click()
lblNome.Enabled = False
cboNome.Enabled = False
lblMes.Enabled = False
cboMes.Enabled = False
lblAno.Enabled = False
cboAno.Enabled = False
lblInicio.Enabled = False
lblFim.Enabled = False
mskInicio.Enabled = False
mskFim.Enabled = False
cmdCal1.Enabled = False
cmdCal2.Enabled = False
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
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskFim.SetFocus
         SelectControl mskFim
      End If
   End If
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
         If mskFim.Enabled = True Then mskFim.SetFocus
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskInicio.SetFocus
         SelectControl mskInicio
      End If
   End If
End Sub

Private Sub cmdExibir_Click()
Dim varINDICE As String       'INDICE PARA ORGANIZAR OS DADOS
Dim varSITUACAO As String
Dim varCriterio As String

If optTodos.Value = True Then
    varCriterio = " "
ElseIf optMensal.Value = True Then
    If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub
    varCriterio = " WHERE  (MONTH(VALIDADE) = " & cboMes.ListIndex + 1 & ") AND (YEAR(VALIDADE) = " & cboAno & ") "
ElseIf optPeriodo.Value = True Then
    If mskInicio.Text = "" Or mskFim.Text = "" Then Exit Sub
    varCriterio = " WHERE (VALIDADE >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (VALIDADE <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103))"
ElseIf optCliente.Value = True Then
    If txtCodClie.Text = "" Then Exit Sub
    varCriterio = " WHERE (COD_CLIENTE = " & txtCodClie.Text & ") "
End If

'Verifica qual o índice a ser utlizado
If optIndiceValidade.Value = True Then
   varINDICE = "Pedidos_Cashback.VALIDADE;"
ElseIf optIndiceCliente.Value = True Then
   varINDICE = "cliente.Nome;"
'ElseIf optCidade.Value = True Then
'   INDICE = "cidade;"
End If

'Verifica a situacao do cliente
If optSituacaoTodos.Value = True And optTodos.Value = True Then
    varSITUACAO = " "
ElseIf optAtivos.Value = True And optTodos.Value = True Then
   varSITUACAO = "WHERE Pedidos_Cashback.ABATIDO = 0 AND Pedidos_Cashback.INVALIDO = 0"
ElseIf optAtivos.Value = True And optTodos.Value = False Then
   varSITUACAO = "AND Pedidos_Cashback.ABATIDO = 0 AND Pedidos_Cashback.INVALIDO = 0"
ElseIf optAbatidos.Value = True And optTodos.Value = True Then
   varSITUACAO = "WHERE Pedidos_Cashback.ABATIDO = 1"
ElseIf optAbatidos.Value = True And optTodos.Value = False Then
   varSITUACAO = "AND Pedidos_Cashback.ABATIDO = 1"
ElseIf optInutilizado.Value = True And optTodos.Value = True Then
    varSITUACAO = "WHERE Pedidos_Cashback.INVALIDO = 1"
ElseIf optInutilizado.Value = True And optTodos.Value = False Then
    varSITUACAO = "AND Pedidos_Cashback.INVALIDO = 1"
End If

 
sSQL = "SELECT Pedidos_Cashback.CODIGO AS varCod, Pedidos_Cashback.COD_PEDIDO AS varCodPed, Pedidos_Cashback.COD_CLIENTE AS varCodCli, cliente.Nome AS varCli, cliente.Celular AS varCelular, Pedidos_Cashback.VALOR_VENDA AS varVenda, Pedidos_Cashback.VALOR_CASHBACK AS varVlrCash, Pedidos_Cashback.VALOR_ABATIDO AS varVlrAbatido, (CASE WHEN Pedidos_Cashback.ABATIDO = 1 THEN 'SIM' ELSE '' END) AS varAbatido, Pedidos_Cashback.DATA_ABATIDO AS varDataAbatido, Pedidos_Cashback.COD_PEDIDOABATIDO AS varCodPedAbatido, Pedidos_Cashback.VALIDADE AS varValidade, (CASE WHEN Pedidos_Cashback.INVALIDO = 1 THEN 'SIM' ELSE '' END) AS varInvalido, Pedidos_Cashback.COD_FUNCIONARIO AS varCodFunc " & _
        "FROM Pedidos_Cashback INNER JOIN cliente ON Pedidos_Cashback.COD_CLIENTE = cliente.CODIGO "
sSQL = sSQL & varCriterio & varSITUACAO & " ORDER BY " & varINDICE

Set r = dbData.OpenRecordset(sSQL)
printSQL = sSQL

FormatarGridConsulta r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdCal1_Click()
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




Private Sub cmdCal2_Click()
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

Private Sub Form_Load()
   Set moCombo = New cComboHelper
   
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub


