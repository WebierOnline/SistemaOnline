VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Consulta_EntradavsSaida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE ENTRADAS"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14955
   Icon            =   "Consulta_EntradavsSaida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   14955
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   60
      ScaleHeight     =   2025
      ScaleWidth      =   14805
      TabIndex        =   5
      ToolTipText     =   "Imprimir"
      Top             =   840
      Width           =   14835
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
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
         Height          =   1875
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   14655
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
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
            Height          =   1515
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2535
            Begin VB.ComboBox cboIndice 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   1080
               Width           =   2175
            End
            Begin VB.ComboBox cboCriterioPrinc 
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Organizar por:"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   840
               Width           =   990
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criterio"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Frame frmFiltro1 
            BackColor       =   &H00E0E0E0&
            Height          =   1500
            Left            =   2700
            TabIndex        =   7
            Top             =   240
            Width           =   11895
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   1500
               Sorted          =   -1  'True
               TabIndex        =   12
               Top             =   480
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMes 
               Height          =   315
               Left            =   120
               TabIndex        =   11
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboCategoria 
               Height          =   315
               Left            =   120
               TabIndex        =   10
               Top             =   1080
               Visible         =   0   'False
               Width           =   4125
            End
            Begin VB.ComboBox cboDescricao 
               Height          =   315
               Left            =   120
               TabIndex        =   9
               Top             =   1080
               Visible         =   0   'False
               Width           =   6615
            End
            Begin VB.TextBox txtCodBarra 
               Height          =   315
               Left            =   120
               TabIndex        =   8
               Top             =   1080
               Visible         =   0   'False
               Width           =   2355
            End
            Begin ChamaleonBtn.chameleonButton cmdCalendario1 
               Height          =   315
               Left            =   1080
               TabIndex        =   25
               Tag             =   "Calendario"
               Top             =   480
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
               MICON           =   "Consulta_EntradavsSaida.frx":23D2
               PICN            =   "Consulta_EntradavsSaida.frx":23EE
               PICH            =   "Consulta_EntradavsSaida.frx":4741
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
               TabIndex        =   26
               Tag             =   "Calendario"
               Top             =   480
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
               BCOL            =   14737632
               BCOLO           =   14737632
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Consulta_EntradavsSaida.frx":6A94
               PICN            =   "Consulta_EntradavsSaida.frx":6AB0
               PICH            =   "Consulta_EntradavsSaida.frx":8E03
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
               TabIndex        =   27
               Top             =   480
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
               TabIndex        =   28
               Top             =   480
               Visible         =   0   'False
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdImprimir 
               Height          =   315
               Left            =   4320
               TabIndex        =   31
               Top             =   480
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
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
               MICON           =   "Consulta_EntradavsSaida.frx":B156
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
               Height          =   315
               Left            =   2820
               TabIndex        =   32
               Top             =   480
               Width           =   1455
               _ExtentX        =   2566
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
               MICON           =   "Consulta_EntradavsSaida.frx":B172
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1500
               TabIndex        =   17
               Top             =   240
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblCategoria 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Categoria:"
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   840
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label lblDescricao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   840
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblCodBarra 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cod. Barra:"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   840
               Visible         =   0   'False
               Width           =   855
            End
         End
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6840
      Picture         =   "Consulta_EntradavsSaida.frx":B18E
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   5400
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
      ScaleWidth      =   14805
      TabIndex        =   0
      Top             =   60
      Width           =   14835
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA DE ESTOQUE - ENTRADAS VS SAÍDAS"
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
         Width           =   7605
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   240
         Picture         =   "Consulta_EntradavsSaida.frx":C1C6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   9270
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22040
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "22:29"
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
      TabIndex        =   18
      Top             =   2940
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   10504
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirEntradas 
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   8940
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "EXIBIR ENTRADAS"
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
      MICON           =   "Consulta_EntradavsSaida.frx":12A0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirVendas 
      Height          =   255
      Left            =   2100
      TabIndex        =   29
      Top             =   8940
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "EXIBIR VENDAS"
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
      MICON           =   "Consulta_EntradavsSaida.frx":12A28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirSaidas 
      Height          =   255
      Left            =   4140
      TabIndex        =   30
      Top             =   8940
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "EXIBIR SAÍDAS"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Consulta_EntradavsSaida.frx":12A44
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   8940
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Consulta_EntradavsSaida"
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
      .Cols = 10
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 600
      .ColWidth(2) = 7000
      .ColWidth(3) = 1200
      .ColWidth(4) = 1100
      .ColWidth(5) = 1100
      .ColWidth(6) = 1100
      .ColWidth(7) = 1100
      .ColWidth(8) = 1100
      .ColWidth(9) = 0
      
      .TextMatrix(0, 1) = "CÓD."
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "ENTRADAS"
      .TextMatrix(0, 4) = "VENDAS"
      .TextMatrix(0, 5) = "SAÍDAS"
      .TextMatrix(0, 6) = "REMOÇÃO"
      .TextMatrix(0, 7) = "TOTAL"
      .TextMatrix(0, 8) = "ESTOQUE"
      .TextMatrix(0, 9) = "CÓD. BARRA"
      
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
            .TextMatrix(.rows - 1, 1) = ValidateNull(rTabela("CODIGO"))
            .TextMatrix(.rows - 1, 2) = rTabela("DESCRICAO") & " /  " & rTabela("FABRICANTE")
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("vEntradas"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("vVendas"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("vSaidas"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("vRemocao"))
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("vEstoqueCalculado"))
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("Quant_Estoque"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("COD_BARRA"))
            
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
   
   'lblQtda.Caption = Format(SomaGrid(Grid, 3), ocPESO)
picAguarde.Visible = False
End Sub
Private Sub LimparObjetos_Consulta()
cboMes.Text = ""
cboAno.Text = ""
cboCategoria.Text = ""
End Sub

Private Sub PreencherIndice()
cboIndice.AddItem "DESCRIÇÃO"
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

Private Sub cboCategoria_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboCategoria.Clear
   
   sSQL = "SELECT categoria FROM produtos GROUP BY categoria;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCategoria.AddItem ValidateNull(r("categoria"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCategoria
End Sub

Private Sub cboCriterioPrinc_Click()
cboCriterioPrinc_LostFocus
End Sub

Private Sub cboCriterioPrinc_GotFocus()
cboCriterioPrinc.Clear
cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "CÓD.BARRA"
cboCriterioPrinc.AddItem "DESCRIÇÃO"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "PERÍODO"
cboCriterioPrinc.AddItem "MENSAL/CÓD.BARRA"
cboCriterioPrinc.AddItem "MENSAL/DESCRIÇÃO"
cboCriterioPrinc.AddItem "MENSAL/CATEGORIA"
cboCriterioPrinc.AddItem "PERÍODO/CÓD.BARRA"
cboCriterioPrinc.AddItem "PERÍODO/DESCRIÇÃO"
cboCriterioPrinc.AddItem "PERÍODO/CATEGORIA"
moCombo.AttachTo cboCriterioPrinc
End Sub


Private Sub cboCriterioPrinc_LostFocus()
If cboCriterioPrinc.Text = "TODOS" Then
   lblMes.Visible = False
   lblAno.Visible = False
   cmdCalendario1.Visible = False
   cmdCalendario2.Visible = False
   mskInicio.Visible = False
   mskFim.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCodBarra.Visible = False
   lblDescricao.Visible = False
   lblCategoria.Visible = False
   txtCodBarra.Visible = False
   cboDescricao.Visible = False
   cboCategoria.Visible = False
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
   lblCodBarra.Visible = False
   lblDescricao.Visible = False
   lblCategoria.Visible = False
   txtCodBarra.Visible = False
   cboDescricao.Visible = False
   cboCategoria.Visible = False
   If cboMes.Visible = True Then cboMes.SetFocus
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
   lblCodBarra.Visible = False
   lblDescricao.Visible = False
   lblCategoria.Visible = False
   txtCodBarra.Visible = False
   cboDescricao.Visible = False
   cboCategoria.Visible = False
   mskInicio.SetFocus
ElseIf cboCriterioPrinc.Text = "MENSAL/CÓD.BARRA" Then
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
   lblCodBarra.Visible = True
   lblDescricao.Visible = False
   lblCategoria.Visible = False
   txtCodBarra.Visible = True
   cboDescricao.Visible = False
   cboCategoria.Visible = False
   txtCodBarra.SetFocus
ElseIf cboCriterioPrinc.Text = "MENSAL/DESCRIÇÃO" Then
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
   lblCodBarra.Visible = False
   lblDescricao.Visible = True
   lblCategoria.Visible = False
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   cboCategoria.Visible = False
   cboDescricao.SetFocus
ElseIf cboCriterioPrinc.Text = "MENSAL/CATEGORIA" Then
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
   lblCodBarra.Visible = False
   lblDescricao.Visible = False
   lblCategoria.Visible = True
   txtCodBarra.Visible = False
   cboDescricao.Visible = False
   cboCategoria.Visible = True
   cboCategoria.SetFocus
ElseIf cboCriterioPrinc.Text = "CÓD.BARRA" Then
   lblMes.Visible = False
   lblAno.Visible = False
   cmdCalendario1.Visible = False
   cmdCalendario2.Visible = False
   mskInicio.Visible = False
   mskFim.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCodBarra.Visible = True
   lblDescricao.Visible = False
   lblCategoria.Visible = False
   txtCodBarra.Visible = True
   cboDescricao.Visible = False
   cboCategoria.Visible = False
   txtCodBarra.SetFocus
ElseIf cboCriterioPrinc.Text = "DESCRIÇÃO" Then
   lblMes.Visible = False
   lblAno.Visible = False
   cmdCalendario1.Visible = False
   cmdCalendario2.Visible = False
   mskInicio.Visible = False
   mskFim.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCodBarra.Visible = False
   lblDescricao.Visible = True
   lblCategoria.Visible = False
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   cboCategoria.Visible = False
   cboDescricao.SetFocus
ElseIf cboCriterioPrinc.Text = "PERÍODO/CÓD.BARRA" Then
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
   lblCodBarra.Visible = True
   lblDescricao.Visible = False
   lblCategoria.Visible = False
   txtCodBarra.Visible = True
   cboDescricao.Visible = False
   cboCategoria.Visible = False
   txtCodBarra.SetFocus
ElseIf cboCriterioPrinc.Text = "PERÍODO/DESCRIÇÃO" Then
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
   lblCodBarra.Visible = False
   lblDescricao.Visible = True
   lblCategoria.Visible = False
   txtCodBarra.Visible = False
   cboDescricao.Visible = True
   cboCategoria.Visible = False
   cboDescricao.SetFocus
ElseIf cboCriterioPrinc.Text = "PERÍODO/CATEGORIA" Then
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
   lblCodBarra.Visible = False
   lblDescricao.Visible = False
   lblCategoria.Visible = True
   txtCodBarra.Visible = False
   cboDescricao.Visible = False
   cboCategoria.Visible = True
   cboCategoria.SetFocus

Else
End If

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

sSQL = "SELECT DISTINCT descricao FROM produtos ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboDescricao.AddItem r("descricao")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboDescricao
End Sub

Private Sub cboIndice_GotFocus()
cboIndice.Clear
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
If cboIndice.Text = "DESCRIÇÃO" Then
   INDICE = "C.DESCRICAO;"
Else
   INDICE = "C.DESCRICAO;"
End If

Dim vCMProd_Quant As String
Dim vCMPed_Itens As String
Dim vCMSaida As String
Dim vCMProd_Quant1 As String
Dim vCMPed_Itens1 As String
Dim vCMSaida1 As String
Dim vCMProd_Quant3 As String
Dim vCriterio As String

If Left(cboCriterioPrinc.Text, 6) = "MENSAL" Then
    vCMProd_Quant1 = " AND (MONTH(Produtos_Quant_1.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant_1.data) = " & cboAno & ") "
    vCMPed_Itens1 = " AND (MONTH(pedidos_itens_1.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens_1.data) = " & cboAno & ") "
    vCMSaida1 = " AND (MONTH(produtos_saida_1.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(produtos_saida_1.data) = " & cboAno & ") "
    vCMProd_Quant = " AND (MONTH(Produtos_Quant.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant.data) = " & cboAno & ") "
    vCMPed_Itens = " AND (MONTH(pedidos_itens.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos_itens.data) = " & cboAno & ") "
    vCMSaida = " AND (MONTH(produtos_saida.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(produtos_saida.data) = " & cboAno & ") "
    vCMProd_Quant3 = " AND (MONTH(Produtos_Quant_3.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant_3.data) = " & cboAno & ") "
    
    If cboCriterioPrinc.Text = "MENSAL/CÓD.BARRA" Then
        vCriterio = " AND (c.cod_barra = '" & txtCodBarra.Text & "') "
    ElseIf cboCriterioPrinc.Text = "MENSAL/DESCRIÇÃO" Then
        vCriterio = " AND (C.descricao = '" & cboDescricao.Text & "') "
    ElseIf cboCriterioPrinc.Text = "MENSAL/CATEGORIA" Then
        vCriterio = " AND (C.categoria = '" & cboCategoria.Text & "') "
    Else
        vCriterio = " "
    End If

ElseIf Left(cboCriterioPrinc.Text, 7) = "PERÍODO" Then
    vCMProd_Quant1 = " AND (Produtos_Quant_1.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant_1.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    vCMPed_Itens1 = " AND (pedidos_itens_1.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens_1.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    vCMSaida1 = " AND (produtos_saida_1.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (produtos_saida_1.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    vCMProd_Quant = " AND (Produtos_Quant.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    vCMPed_Itens = " AND (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    vCMSaida = " AND (produtos_saida.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (produtos_saida.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "
    vCMProd_Quant3 = " AND (Produtos_Quant_3.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant_3.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "

    If cboCriterioPrinc.Text = "PERÍODO/CÓD.BARRA" Then
        vCriterio = " AND (c.cod_barra = '" & txtCodBarra.Text & "') "
    ElseIf cboCriterioPrinc.Text = "PERÍODO/DESCRIÇÃO" Then
        vCriterio = " AND (C.descricao = '" & cboDescricao.Text & "') "
    ElseIf cboCriterioPrinc.Text = "PERÍODO/CATEGORIA" Then
        vCriterio = " AND (C.categoria = '" & cboCategoria.Text & "') "
    Else
        vCriterio = " "
    End If

Else
    vCMProd_Quant = " "
    vCMPed_Itens = " "
    vCMSaida = " "
    vCMProd_Quant1 = " "
    vCMPed_Itens1 = " "
    vCMSaida1 = " "
    vCMProd_Quant3 = " "
    vCriterio = " "
End If

sSQL = "SELECT DISTINCT C.CODIGO, C.DESCRICAO, C.FABRICANTE, C.COD_BARRA, " & _
       "(SELECT isnull(SUM(QUANT), 0) AS Expr1 FROM Produtos_Quant WHERE (COD_PRODUTO = C.CODIGO) AND (TIPO <> 'REMOÇÃO') " & vCMProd_Quant & " " & vCriterio & ") AS vEntradas, " & _
       "(SELECT isnull(SUM(pedidos_itens.QUANTIDADE), 0) AS Expr1 FROM pedidos_itens INNER JOIN pedidos ON pedidos_itens.COD_PEDIDO = pedidos.COD_PEDIDO WHERE (pedidos_itens.cancelado = 0) AND (pedidos.TIPO_PEDIDO <> 'ORÇAMENTO') AND (pedidos_itens.COD_PRODUTO = C.CODIGO) " & vCMPed_Itens & " " & vCriterio & ") AS vVendas, " & _
       "(SELECT ISNULL(SUM(saida), 0) AS Expr1 FROM produtos_saida WHERE (cod_produto = C.CODIGO) AND (EXCLUIDO = 0) " & vCMSaida & " " & vCriterio & ") AS vSaidas, " & _
       "(SELECT ISNULL(SUM(QUANT), 0) AS Expr1 FROM Produtos_Quant AS Produtos_Quant_3 WHERE (COD_PRODUTO = C.CODIGO) AND (TIPO = 'REMOÇÃO') AND (FORMA = 'AJUSTE') " & vCMProd_Quant3 & " " & vCriterio & ") AS vRemocao, " & _
       "(SELECT SUM(QUANT) AS Expr1 FROM Produtos_Quant AS Produtos_Quant_1 WHERE (COD_PRODUTO = C.CODIGO) AND (TIPO <> 'REMOÇÃO') " & vCMProd_Quant1 & " " & vCriterio & ") -  " & _
       "((SELECT SUM(pedidos_itens_1.QUANTIDADE) AS Expr1 FROM pedidos_itens AS pedidos_itens_1 INNER JOIN pedidos AS pedidos_1 ON pedidos_itens_1.COD_PEDIDO = pedidos_1.COD_PEDIDO WHERE (pedidos_itens_1.cancelado = 0) AND (pedidos_1.TIPO_PEDIDO <> 'ORÇAMENTO') AND (pedidos_itens_1.COD_PRODUTO = C.CODIGO) " & vCMPed_Itens1 & " " & vCriterio & ") + (SELECT ISNULL(SUM(saida), 0) AS Expr1 FROM produtos_saida AS produtos_saida_1 WHERE (cod_produto = C.CODIGO) AND (EXCLUIDO = 0) " & vCMSaida1 & " " & vCriterio & ") + (SELECT ISNULL(SUM(QUANT), 0) AS Expr1 FROM Produtos_Quant AS Produtos_Quant_3 WHERE (COD_PRODUTO = C.CODIGO) AND (TIPO = 'REMOÇÃO') AND (FORMA = 'AJUSTE') " & vCMProd_Quant3 & " " & vCriterio & ")) AS vEstoqueCalculado,  " & _
       "C.QUANT_ESTOQUE " & _
       "FROM produtos AS C INNER JOIN Produtos_Quant AS Produtos_Quant_2 ON C.CODIGO = Produtos_Quant_2.COD_PRODUTO " & _
       "WHERE (Produtos_Quant_2.TIPO <> 'REMOÇÃO') "
       
If cboCriterioPrinc.Text = "TODOS" Then

ElseIf cboCriterioPrinc.Text = "CÓD.BARRA" Then
    If txtCodBarra.Text = "" Then Exit Sub
    sSQL = sSQL & " AND (c.cod_barra = '" & txtCodBarra.Text & "') "

ElseIf cboCriterioPrinc.Text = "DESCRIÇÃO" Then
    If cboDescricao.Text = "" Then Exit Sub
    sSQL = sSQL & " AND (C.descricao = '" & cboDescricao.Text & "') "
    
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sSQL = sSQL & " AND (MONTH(Produtos_Quant_2.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant_2.data) = " & cboAno & ")"

ElseIf cboCriterioPrinc.Text = "PERÍODO" Then
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sSQL = sSQL & " AND (Produtos_Quant_2.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant_2.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "

ElseIf cboCriterioPrinc.Text = "MENSAL/CÓD.BARRA" Then
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sSQL = sSQL & " AND (c.cod_barra = '" & txtCodBarra.Text & "') AND (MONTH(Produtos_Quant_2.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant_2.data) = " & cboAno & ") "

ElseIf cboCriterioPrinc.Text = "MENSAL/DESCRIÇÃO" Then
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sSQL = sSQL & " AND (C.descricao = '" & cboDescricao.Text & "') AND (MONTH(Produtos_Quant_2.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant_2.data) = " & cboAno & ") "

ElseIf cboCriterioPrinc.Text = "MENSAL/CATEGORIA" Then
    If cboAno.Text = "" Or cboMes.Text = "" Then Exit Sub
    sSQL = sSQL & " AND (C.categoria = '" & cboCategoria.Text & "') AND (MONTH(Produtos_Quant_2.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(Produtos_Quant_2.data) = " & cboAno & ") "

ElseIf cboCriterioPrinc.Text = "PERÍODO/CÓD.BARRA" Then
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sSQL = sSQL & " AND (c.cod_barra = '" & txtCodBarra.Text & "') AND (Produtos_Quant_2.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant_2.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "

ElseIf cboCriterioPrinc.Text = "PERÍODO/DESCRIÇÃO" Then
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sSQL = sSQL & " AND (C.descricao = '" & cboDescricao.Text & "') AND (Produtos_Quant_2.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant_2.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "

ElseIf cboCriterioPrinc.Text = "PERÍODO/CATEGORIA" Then
    If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub
    sSQL = sSQL & " AND (C.categoria = '" & cboCategoria.Text & "') AND (Produtos_Quant_2.data >= CONVERT(DATETIME, '" & Format(mskInicio.Text, ocDATA) & "', 103)) AND (Produtos_Quant_2.data <= CONVERT(DATETIME, '" & Format(mskFim.Text, ocDATA) & "', 103)) "

End If

sSQL = sSQL & "ORDER BY " & INDICE
'Debug.Print sSQL

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
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
      Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.loadPedidos (Grid.TextMatrix(Grid.Row, 1))
      'Me.Hide
      Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.Show
End If
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

Private Sub cmdExibirVendas_Click()
'Vendas_Consulta_PorProdutos

If Grid.Col = 0 Then Exit Sub
'If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
'      Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.loadPedidos (Grid.TextMatrix(Grid.Row, 1))
'      'Me.Hide
'      Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.Show
'End If


Vendas_Consulta_PorProdutos.cboCriterioPrinc.Text = "TODOS"
Vendas_Consulta_PorProdutos.cboTipo.Text = "POR PRODUTOS"
Vendas_Consulta_PorProdutos.cboCriterioSec.Text = "CÓD. BARRA"
Vendas_Consulta_PorProdutos.txtCodBarra.Text = Grid.TextMatrix(Grid.Row, 9)
Vendas_Consulta_PorProdutos.cmdLocalizar_Click
Vendas_Consulta_PorProdutos.Show
End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset

Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide    'ver depois como nao exibir

Set r = dbData.OpenRecordset(printSQL)
Set REL_Cons_Entrada_ProdAgrupado.Relatorio.Recordset = r

If cboCriterioPrinc.Text = "TODOS" Then
    REL_Cons_Entrada_ProdAgrupado.rfCons1.Caption = "TODOS"
    REL_Cons_Entrada_ProdAgrupado.rfCons3.Caption = ""
ElseIf cboCriterioPrinc.Text = "MENSAL" Then
    REL_Cons_Entrada_ProdAgrupado.rfCons1.Caption = "MENSAL"
    REL_Cons_Entrada_ProdAgrupado.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
ElseIf cboCriterioPrinc.Text = "MENSAL/CÓD.BARRA" Then
    REL_Cons_Entrada_ProdAgrupado.rfCons1.Caption = "MENSAL/CÓD.BARRA"
    REL_Cons_Entrada_ProdAgrupado.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
    REL_Cons_Entrada_ProdAgrupado.rfCons2.Caption = "Cód. Barra = " & txtCodBarra.Text & ""
ElseIf cboCriterioPrinc.Text = "MENSAL/DESCRIÇÃO" Then
    REL_Cons_Entrada_ProdAgrupado.rfCons1.Caption = "MENSAL/DESCRIÇÃO"
    REL_Cons_Entrada_ProdAgrupado.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
    REL_Cons_Entrada_ProdAgrupado.rfCons2.Caption = "DESCRIÇÃO = " & cboDescricao.Text & ""
ElseIf cboCriterioPrinc.Text = "MENSAL/CATEGORIA" Then
    REL_Cons_Entrada_ProdAgrupado.rfCons1.Caption = "MENSAL/CATEGORIA"
    REL_Cons_Entrada_ProdAgrupado.rfCons3.Caption = "Mês/Ano = " & cboMes.Text & "/" & cboAno.Text
    REL_Cons_Entrada_ProdAgrupado.rfCons2.Caption = "CATEGORIA = " & cboCategoria.Text & ""
End If

'REL_Cons_Entrada_ProdAgrupado.dfQuant.Caption = lblQtda.Caption

'REL_Cons_Entrada_ProdAgrupado.Relatorio.NomeImpressora = var_Impressora
REL_Cons_Entrada_ProdAgrupado.Relatorio.Ativar
Unload REL_Cons_Entrada_ProdAgrupado

Me.Show 1   'ver depois como nao exibir
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

PreencherCriterio
cboCriterioPrinc.ListIndex = 3

PreencherIndice
cboIndice.ListIndex = 0

'cboMes.Text = Format(Date, "mmmm")
'cboAno.Text = Year(Date)

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
Set moCombo = New cComboHelper
End Sub
Private Sub PreencherCriterio()
cboCriterioPrinc.AddItem "TODOS"
cboCriterioPrinc.AddItem "CÓD.BARRA"
cboCriterioPrinc.AddItem "DESCRIÇÃO"
cboCriterioPrinc.AddItem "MENSAL"
cboCriterioPrinc.AddItem "PERÍODO"
cboCriterioPrinc.AddItem "MENSAL/CÓD.BARRA"
cboCriterioPrinc.AddItem "MENSAL/DESCRIÇÃO"
cboCriterioPrinc.AddItem "MENSAL/CATEGORIA"
cboCriterioPrinc.AddItem "PERÍODO/CÓD.BARRA"
cboCriterioPrinc.AddItem "PERÍODO/DESCRIÇÃO"
cboCriterioPrinc.AddItem "PERÍODO/CATEGORIA"
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

