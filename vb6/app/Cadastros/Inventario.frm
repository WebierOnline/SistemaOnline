VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Inventario_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVENTÁRIO"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "Inventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6906
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "Inventario.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalvar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmCadastro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Inventario.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdExibir"
      Tab(1).Control(1)=   "cmdExcluir"
      Tab(1).Control(2)=   "cmdDetalhar"
      Tab(1).Control(3)=   "Grid"
      Tab(1).ControlCount=   4
      Begin VB.Frame frmCadastro 
         Caption         =   "Inventário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3315
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6255
         Begin VB.ComboBox cboAno 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3240
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cboPeriodo 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   14
            Top             =   1080
            Width           =   1515
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1515
         End
         Begin VB.ComboBox cboFuncionario 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   5955
         End
         Begin VB.TextBox txtCodFuncionario 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   5400
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   735
         End
         Begin ChamaleonBtn.chameleonButton cmdCal1 
            Height          =   315
            Left            =   5160
            TabIndex        =   9
            Tag             =   "Calendario"
            Top             =   1080
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
            MICON           =   "Inventario.frx":240A
            PICN            =   "Inventario.frx":2426
            PICH            =   "Inventario.frx":4779
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   4260
            TabIndex        =   10
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            Caption         =   "Status"
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
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   555
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
            Left            =   3240
            TabIndex        =   17
            Top             =   840
            Width           =   345
         End
         Begin VB.Label lblPeriodo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Período"
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
            Left            =   1680
            TabIndex        =   16
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   390
         End
         Begin VB.Label lvlData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4260
            TabIndex        =   11
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionário"
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
            TabIndex        =   8
            Top             =   240
            Width           =   990
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   6420
         TabIndex        =   18
         Top             =   1740
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Cancelar"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inventario.frx":6ACC
         PICN            =   "Inventario.frx":6AE8
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
         Height          =   615
         Left            =   6420
         TabIndex        =   19
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Salvar"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inventario.frx":887A
         PICN            =   "Inventario.frx":8896
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   615
         Left            =   6420
         TabIndex        =   20
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inventario.frx":A628
         PICN            =   "Inventario.frx":A644
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   21
         Top             =   780
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5424
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin ChamaleonBtn.chameleonButton cmdDetalhar 
         Height          =   315
         Left            =   -72660
         TabIndex        =   23
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Detalhar"
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
         MICON           =   "Inventario.frx":C3D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluir 
         Height          =   315
         Left            =   -70440
         TabIndex        =   24
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "Inventario.frx":C3F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   315
         Left            =   -74880
         TabIndex        =   25
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Exibir"
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
         MICON           =   "Inventario.frx":C40E
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
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   8745
      TabIndex        =   1
      Top             =   60
      Width           =   8775
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Inventario.frx":C42A
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
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
         Left            =   1020
         TabIndex        =   2
         Top             =   180
         Width           =   1860
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   4875
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11404
            Text            =   "Desenv.: Online.Info Sistemas"
            TextSave        =   "Desenv.: Online.Info Sistemas"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "06:30"
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
End
Attribute VB_Name = "Inventario_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Dim bRet As Boolean
Dim vCod As Integer
Dim vQuantLinhas As Integer




Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer, x As Integer

With Grid
   .Clear
   .Cols = 9
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 500
   .ColWidth(2) = 1000
   .ColWidth(3) = 500
   .ColWidth(4) = 1000
   .ColWidth(5) = 1000
   .ColWidth(6) = 1000
   .ColWidth(7) = 1000
   .ColWidth(8) = 1000
   
   For x = 0 To .Cols - 1
      .Col = x
      .Row = 0
      .CellFontBold = True
   Next
   'IdInventario, Data, IDFuncionario, AnoInventario, PeriodoInventarioTipo, PeriodoInventario, DataInventario,
   .TextMatrix(0, 1) = "ID"
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "FUNC"
   .TextMatrix(0, 4) = "ANO"
   .TextMatrix(0, 5) = "TIPO"
   .TextMatrix(0, 6) = "PERÍODO"
   .TextMatrix(0, 7) = "DATA"
   .TextMatrix(0, 8) = "FECHADO"
   .Redraw = False
   
   i = 1
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("IdInventario")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("DATA"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 3) = rTabela("IDFuncionario")
         .TextMatrix(.rows - 1, 4) = rTabela("AnoInventario")
         .TextMatrix(.rows - 1, 5) = rTabela("PeriodoInventarioTipo")
         .TextMatrix(.rows - 1, 6) = rTabela("PeriodoInventario")
         .TextMatrix(.rows - 1, 7) = Format(rTabela("DataInventario"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 8) = rTabela("VFECHADO")
         rTabela.MoveNext
         
         .rows = .rows + 1
         i = i + 1
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
End Sub

Private Sub Limpar_Objetos()
txtCodigo.Text = ""
cboFuncionario.Text = ""
txtCodFuncionario.Text = ""
cboTipo.Text = ""
cboPeriodo.Text = ""
cboAno.Text = ""
mskData.Mask = ""
mskData.Text = ""
lblStatus.Caption = "Status"
End Sub

Private Sub cboAno_GotFocus()
'moCombo.AttachTo cboAno
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAno.Clear

iAno = Year(Date)
' O último ano será o anterior ao atual
LastYear = iAno - 1
' O primeiro ano será 5 anos atrás em relação ao LastYear
FirstYear = iAno - 5

For i = FirstYear To LastYear
   cboAno.AddItem CStr(i)
Next

' Se o moCombo for um componente de subclasse ou skin, mantenha a linha abaixo
' moCombo.AttachTo cboAno

' Opcional: Deixar o ano mais recente (ano passado) já selecionado
If cboAno.ListCount > 0 Then
    cboAno.ListIndex = cboAno.ListCount - 1
End If
End Sub

Private Sub cboPeriodo_GotFocus()
cboPeriodo.Clear

If cboTipo.Text = "Anual" Then

ElseIf cboTipo.Text = "Semestral" Then
    cboPeriodo.AddItem "1"
    cboPeriodo.AddItem "2"
ElseIf cboTipo.Text = "Trimestral" Then
    cboPeriodo.AddItem "1"
    cboPeriodo.AddItem "2"
    cboPeriodo.AddItem "3"
    cboPeriodo.AddItem "4"
ElseIf cboTipo.Text = "Bimestral" Then
    cboPeriodo.AddItem "1"
    cboPeriodo.AddItem "2"
    cboPeriodo.AddItem "3"
    cboPeriodo.AddItem "4"
    cboPeriodo.AddItem "5"
    cboPeriodo.AddItem "6"
ElseIf cboTipo.Text = "Mensal" Then
    cboPeriodo.AddItem "1"
    cboPeriodo.AddItem "2"
    cboPeriodo.AddItem "3"
    cboPeriodo.AddItem "4"
    cboPeriodo.AddItem "5"
    cboPeriodo.AddItem "6"
    cboPeriodo.AddItem "7"
    cboPeriodo.AddItem "8"
    cboPeriodo.AddItem "9"
    cboPeriodo.AddItem "10"
    cboPeriodo.AddItem "11"
    cboPeriodo.AddItem "12"
ElseIf cboTipo.Text = "Dia" Then

End If
moCombo.AttachTo cboPeriodo
End Sub


Private Sub cboTipo_Click()
cboTipo_LostFocus
End Sub

Private Sub cboTipo_GotFocus()
cboTipo.Clear
cboTipo.AddItem "Anual"
'cboTipo.AddItem "Semestral"
'cboTipo.AddItem "Trimestral"
'cboTipo.AddItem "Bimestral"
'cboTipo.AddItem "Mensal"
'cboTipo.AddItem "Dia"
moCombo.AttachTo cboTipo
End Sub


Private Sub cboTipo_LostFocus()
If cboTipo.Text = "Anual" Then
    lblPeriodo.Enabled = False
    cboPeriodo.Enabled = False
    lblAno.Enabled = True
    cboAno.Enabled = True
    lvlData.Enabled = False
    mskData.Enabled = False
    cmdCal1.Enabled = False
ElseIf cboTipo.Text = "Semestral" Then
    lblPeriodo.Enabled = True
    cboPeriodo.Enabled = True
    lblAno.Enabled = True
    cboAno.Enabled = True
    lvlData.Enabled = False
    mskData.Enabled = False
    cmdCal1.Enabled = False
ElseIf cboTipo.Text = "Trimestral" Then
    lblPeriodo.Enabled = True
    cboPeriodo.Enabled = True
    lblAno.Enabled = True
    cboAno.Enabled = True
    lvlData.Enabled = False
    mskData.Enabled = False
    cmdCal1.Enabled = False
ElseIf cboTipo.Text = "Bimestral" Then
    lblPeriodo.Enabled = True
    cboPeriodo.Enabled = True
    lblAno.Enabled = True
    cboAno.Enabled = True
    lvlData.Enabled = False
    mskData.Enabled = False
    cmdCal1.Enabled = False
ElseIf cboTipo.Text = "Mensal" Then
    lblPeriodo.Enabled = True
    cboPeriodo.Enabled = True
    lblAno.Enabled = True
    cboAno.Enabled = True
    lvlData.Enabled = False
    mskData.Enabled = False
    cmdCal1.Enabled = False
ElseIf cboTipo.Text = "Dia" Then
    lblPeriodo.Enabled = True
    cboPeriodo.Enabled = True
    lblAno.Enabled = False
    cboAno.Enabled = False
    lvlData.Enabled = True
    mskData.Enabled = True
    cmdCal1.Enabled = True
End If
End Sub


Private Sub cmdCancelar_Click()
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = False
cmdNovo.Enabled = True
Limpar_Objetos
Form_Load
End Sub

Private Sub cmdDetalhar_Click()
vQuantLinhas = Grid.rows - 1
If vQuantLinhas < 1 Then Exit Sub
vCod = (Grid.TextMatrix(Grid.Row, 1))
Load Inventario_Detalhamento
Inventario_Detalhamento.LerInventario vCod
Inventario_Detalhamento.Show 1
End Sub

Private Sub cmdExcluir_Click()
vQuantLinhas = Grid.rows - 1
If vQuantLinhas < 1 Then Exit Sub

If ShowMsg("Tem certeza que deseja excluir esse inventário?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

vCod = (Grid.TextMatrix(Grid.Row, 1))

sSQL = "DELETE FROM TbInventarios WHERE (IdInventario = " & vCod & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

sSQL = "DELETE FROM TbInventariosItens WHERE (IdInventario = " & vCod & ");"
bRet = dbData.Execute(sSQL)

If Not bRet Then
   ShowMsg "Não foi possível excluir o registro.", vbCritical
   Exit Sub
End If

Limpar_Objetos
Form_Load
End Sub

Private Sub cmdExibir_Click()
'vQuantLinhas = Grid.rows - 1
'If vQuantLinhas < 1 Then Exit Sub
MostrarGrid
End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
'cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = False
frmCadastro.Enabled = True
cboFuncionario.SetFocus
End Sub

Private Sub cmdSalvar_Click()
Dim ComandoSQL As String, filtroSQL As String
Dim vFilial As Integer
vFilial = 1
  
  If Len(cboAno) < 4 Then
    MsgBox "Ano Invalido", vbExclamation, "ATENÇÃO"
    Exit Sub
  End If
  
  If Vazio(txtCodFuncionario.Text) Or txtCodFuncionario.Text = 0 Then
    MsgBox "Responsável Inválido", vbExclamation, "ATENÇÃO"
    Exit Sub
  End If
  
  'btOK.Enabled = False
  DoEvents
  
  'Anual|Semestral|Trimestral|Bimestral|Mensal|Dia
  Select Case cboTipo.Text
     Case "Anual"
        filtroSQL = "PeriodoInventarioTipo = 'Anual' AND AnoInventario = " & cboAno.Text
     Case "Semestral"
        If Int(cboPeriodo.Text) > 2 Then
           MsgBox "Período informado inválido para SEMESTRAL!", vbExclamation + vbOKOnly
           'btOK.Enabled = True
           DoEvents
           Exit Sub
        End If
        filtroSQL = "PeriodoInventarioTipo = '" & cboTipo.Text & "' AND AnoInventario = " & cboAno.Text & " AND PeriodoInventario = '" & cboPeriodo.Text & "'"
     Case "Trimestral"
        If Int(cboPeriodo.Text) > 4 Then
           MsgBox "Período informado inválido para SEMESTRAL!", vbExclamation + vbOKOnly
           'btOK.Enabled = True
           DoEvents
           Exit Sub
        End If
        filtroSQL = "PeriodoInventarioTipo = '" & cboTipo.Text & "' AND AnoInventario = " & cboAno.Text & " AND PeriodoInventario = '" & cboPeriodo.Text & "'"
     Case "Bimestral"
        If Int(cboPeriodo.Text) > 6 Then
           MsgBox "Período informado inválido para SEMESTRAL!", vbExclamation + vbOKOnly
           'btOK.Enabled = True
           DoEvents
           Exit Sub
        End If
        filtroSQL = "PeriodoInventarioTipo = '" & cboTipo.Text & "' AND AnoInventario = " & cboAno.Text & " AND PeriodoInventario = '" & cboPeriodo.Text & "'"
     Case "Mensal"
        filtroSQL = "PeriodoInventarioTipo = '" & cboTipo.Text & "' AND AnoInventario = " & cboAno.Text & " AND PeriodoInventario = '" & cboPeriodo.Text & "'"
     Case Else
        If Not IsDate(mskData.Text) Then
           MsgBox "Data Inválida!", vbCritical + vbOKOnly, "SISC"
           'btOK.Enabled = True
           DoEvents
           mskData.SetFocus
           Exit Sub
        End If
        If cboAno <> Year(mskData.Text) Then
           cboAno = Year(mskData.Text)
           cboAno.Text = Year(mskData.Text)
        End If
        filtroSQL = "PeriodoInventarioTipo = 'Dia' AND DataInventario = " & FdtSQL(mskData.Text)
  End Select
  
  filtroSQL = filtroSQL & " AND IdFilial = " & vFilial
  
  'Prepara
  ComandoSQL = "SELECT COUNT(IdInventario) As r FROM TbInventarios WHERE " & filtroSQL
  
  'Verifica
  If SQLExecutaRetorno(ComandoSQL, "r", 0) > 0 Then
      'Pega o Id
      ComandoSQL = "SELECT IdInventario As r FROM TbInventarios WHERE " & filtroSQL
      
      'Chama o inventário gerado
      'frmInvGerad.xIdInvent = SQLExecutaRetorno(ComandoSQL, "r", 0)
      'Load frmInvGerad
      'frmInvGerad.Show vbModal
      'btOK.Enabled = True
      DoEvents
  Else
      lblStatus.Caption = "Aguarde... Gerando Inventário!"
      DoEvents
      
      'Prepara
        If cboTipo.Text = "Anual" Then
            ComandoSQL = "EXEC InventarioGerar " & cboAno.Text & ", " & vFilial & ", '" & cboTipo.Text & "', '', NULL"
        ElseIf cboTipo.Text = "Semestral" Then
        ElseIf cboTipo.Text = "Trimestral" Then
        ElseIf cboTipo.Text = "Bimestral" Then
        ElseIf cboTipo.Text = "Mensal" Then
        ElseIf cboTipo.Text = "Dia" Then
        End If
          
      'ComandoSQL = "EXEC InventarioGerar " & cboAno.Text & ", " & txtCodFuncionario.Text & ", " & vFilial & ", '" & cboTipo.Text & "', '" & cboPeriodo.Text & "', " & FdtSQL(mskData.Text)
                   'EXEC InventarioGerar 2021, 1, 'Anual', '', NULL
      'Executa
      If SQLExecutaTratandoErro(ComandoSQL) Then
         lblStatus.Caption = "Erro ao Gerar o Inventário!"
         'btOK.Enabled = True
         DoEvents
         Exit Sub
      End If
      
      'Chama o inventário gerado
      'btOK.Caption = "OK"
      MsgBox "Inventário gerado!", vbInformation, "Sistema"
      lblStatus.Caption = "Inventário gerado!"
      DoEvents
      'btOK.Enabled = True
      DoEvents
      'btOK.SetFocus
  End If
  
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = False
cmdNovo.Enabled = True
'On Error GoTo TrataErro

'If txtCodigo.Text = "" Or frmCadastro.Text = "" Then Exit Sub

'If Not Inserir_Dados Then
'   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
'   Exit Sub
'End If

'Limpar_Objetos
'Form_Load
'Exit Sub
   
'TrataErro:
'   If Err.Number = 3022 Then
'      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
'      Exit Sub
'   End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
'cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = False
lblPeriodo.Enabled = False
cboPeriodo.Enabled = False
lblAno.Enabled = False
cboAno.Enabled = False
lvlData.Enabled = False
mskData.Enabled = False
cmdCal1.Enabled = False
cmdNovo.Enabled = True
SSTab1.Tab = 0
MostrarGrid
'If Tela_Principal.StatusBar1.Panels(2).Text <> "PROGRAMADOR" Then
'    cmdExcluir.Enabled = False
'Else
'    cmdExcluir.Enabled = True
'End If
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

mskData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub
Private Sub mskData_GotFocus()
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If mskData.Text = "" Or mskData.Text = "__/__/__" Then
      mskData.Mask = ""
      mskData.Text = ""
   Else
      If Not IsDate(mskData.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskData.SetFocus
      End If
   End If
End Sub

Private Sub txtCodFuncionario_Change()
If txtCodFuncionario.Text = "" Then Exit Sub
If txtCodFuncionario.Text = 0 Then Exit Sub

'txtCodFunc.Text = txtCodFuncionario.Text
'txtCodFuncAP.Text = txtCodFuncionario.Text

'If cmdAlterar.Enabled = True Then
   sSQL = "SELECT * FROM funcionario WHERE (codigo = " & txtCodFuncionario.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then cboFuncionario.Text = r("nome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
'End If
End Sub
Private Sub cboFuncionario_GotFocus()
Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboFuncionario.Text
varCodAntes = txtCodFuncionario.Text

cboFuncionario.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE (cargo <> 'mecânico') ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFuncionario.AddItem r("nome")
   cboFuncionario.ItemData(cboFuncionario.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

txtCodFuncionario.Text = varCodAntes
cboFuncionario.Text = varNomeAntes

cboFuncionario.SelStart = 0
cboFuncionario.SelLength = Len(cboFuncionario)
   
   moCombo.AttachTo cboFuncionario
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFuncionario_LostFocus()
   On Error GoTo TrataErro
   
   If cboFuncionario.Text = "" Then txtCodFuncionario.Text = "": Exit Sub
   
   'If cmdAlterar.Enabled = False Then
      If cboFuncionario.ListIndex = -1 Then
         'txtCodFuncionario.Text = ""
         'Exit Sub
      End If
   'End If
   
   txtCodFuncionario = cboFuncionario.ItemData(cboFuncionario.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
'txtCodigo.Text = ""
'txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
'frmCadastro.Text = (Grid.TextMatrix(Grid.Row, 2))
'cmdAlterar.Enabled = True
'cmdExcluir.Enabled = True
'cmdSalvar.Enabled = False
'cmdCancelar.Enabled = False
'frmCadastro.Enabled = True
'SSTab1.Tab = 0
End Sub

Private Sub MostrarGrid()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT  IdInventario, Data, IDFuncionario, AnoInventario, PeriodoInventarioTipo, PeriodoInventario, DataInventario, (CASE WHEN fechado = 1 THEN 'SIM' ELSE 'NÃO' END) as vFechado FROM TbInventarios ORDER BY IdInventario;"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

