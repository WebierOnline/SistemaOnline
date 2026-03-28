VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Ordem_Servicos_Consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDEM DE SERVIÇO"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   Icon            =   "Ordem_Servicos_Consulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Financeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1620
      TabIndex        =   18
      Top             =   7620
      Width           =   1455
      Begin VB.OptionButton optFinTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton OptAbertos 
         Caption         =   "Abertos"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optFechados 
         Caption         =   "Fechados"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   780
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   10845
      TabIndex        =   16
      Top             =   60
      Width           =   10875
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTA - ORDEM DE SERVIÇOS"
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
         TabIndex        =   17
         Top             =   300
         Width           =   5370
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   120
         Picture         =   "Ordem_Servicos_Consulta.frx":23D2
         Top             =   45
         Width           =   900
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6015
      Left            =   60
      TabIndex        =   15
      Top             =   1080
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   10610
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consultar por:"
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
      Left            =   60
      TabIndex        =   0
      Top             =   7380
      Width           =   9495
      Begin VB.Frame Frame2 
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
         Height          =   1575
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1395
         Begin VB.OptionButton optTodos 
            Caption         =   "&Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optData 
            Caption         =   "&Data"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optMes 
            Caption         =   "&Mês"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   900
            Width           =   855
         End
         Begin VB.OptionButton optCliente 
            Caption         =   "&Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame frmLocalizar 
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
         Height          =   1575
         Left            =   3060
         TabIndex        =   22
         Top             =   240
         Width           =   1575
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3420
            ScaleHeight     =   375
            ScaleWidth      =   3075
            TabIndex        =   28
            Top             =   180
            Width           =   3075
         End
         Begin VB.OptionButton optStatusTerminado 
            Caption         =   "&Terminado"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1260
            Width           =   1395
         End
         Begin VB.OptionButton optStatusAguardando 
            Caption         =   "A&guardando"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1020
            Width           =   1395
         End
         Begin VB.OptionButton optStatusExecucao 
            Caption         =   "Em E&xecução"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   780
            Width           =   1395
         End
         Begin VB.OptionButton optStatusComecar 
            Caption         =   "À C&omeçar"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   540
            Width           =   1395
         End
         Begin VB.OptionButton optStatusTodos 
            Caption         =   "&Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1515
         Left            =   4680
         ScaleHeight     =   1455
         ScaleWidth      =   4635
         TabIndex        =   1
         Top             =   300
         Width           =   4695
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            TabIndex        =   35
            Top             =   780
            Visible         =   0   'False
            Width           =   855
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirMes 
            Height          =   315
            Left            =   3720
            TabIndex        =   9
            Top             =   180
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Exibir"
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
            MICON           =   "Ordem_Servicos_Consulta.frx":2902
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboAno 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2340
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   180
            Width           =   1335
         End
         Begin VB.ComboBox cboCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   540
            TabIndex        =   3
            Top             =   1020
            Width           =   3135
         End
         Begin VB.ComboBox cboMes 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Ordem_Servicos_Consulta.frx":291E
            Left            =   540
            List            =   "Ordem_Servicos_Consulta.frx":2920
            TabIndex        =   2
            Top             =   180
            Width           =   1755
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   540
            TabIndex        =   5
            Top             =   600
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirData 
            Height          =   315
            Left            =   3720
            TabIndex        =   10
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Exibir"
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
            MICON           =   "Ordem_Servicos_Consulta.frx":2922
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirCliente 
            Height          =   315
            Left            =   3720
            TabIndex        =   11
            Top             =   1020
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Exibir"
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
            MICON           =   "Ordem_Servicos_Consulta.frx":293E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblCliente 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            Enabled         =   0   'False
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   1020
            Width           =   465
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   390
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Enabled         =   0   'False
            Height          =   315
            Left            =   180
            TabIndex        =   6
            Top             =   180
            Width           =   345
         End
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   615
      Left            =   9600
      TabIndex        =   13
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
      MICON           =   "Ordem_Servicos_Consulta.frx":295A
      PICN            =   "Ordem_Servicos_Consulta.frx":2976
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
      Height          =   615
      Left            =   9600
      TabIndex        =   14
      Top             =   8100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "Ordem_Servicos_Consulta.frx":2C90
      PICN            =   "Ordem_Servicos_Consulta.frx":2CAC
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
      Top             =   9465
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15108
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "10:43"
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
   Begin VB.Label lblQuant 
      Alignment       =   1  'Right Justify
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
      Left            =   10620
      TabIndex        =   12
      Top             =   7140
      Width           =   225
   End
End
Attribute VB_Name = "Ordem_Servicos_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

'Dim TIPO_STATUS As String

Private Sub Montar_Grid()
   Dim SITUACAO As String     'status
   Dim var_STATUS As String   'financeiro
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   If optStatusTodos.Value = True Then
      SITUACAO = ""
   ElseIf optStatusComecar.Value = True Then
      SITUACAO = "AND (OS.status = 'À COMEÇAR') "
   ElseIf optStatusExecucao.Value = True Then
      SITUACAO = "AND (OS.status = 'EM EXECUÇÃO') "
   ElseIf optStatusAguardando.Value = True Then
      SITUACAO = "AND (OS.status = 'AGUARDANDO') "
   ElseIf optStatusTerminado.Value = True Then
      SITUACAO = "AND (OS.status = 'TERMINADO') "
   End If
   
   If optFinTodos.Value = True Then
      var_STATUS = ""
   ElseIf OptAbertos.Value = True Then
      var_STATUS = "AND (status_os = 0)"
   ElseIf optFechados.Value = True Then
      var_STATUS = "AND (status_os = 1)"
   End If
   
   If optTodos.Value = True Then
      sSQL = "SELECT cliente.*, OS.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' ELSE 'ABERTO' END AS var_status_os, OS.* " & _
         "FROM cliente INNER JOIN OS ON cliente.codigo = OS.cod_cliente WHERE (cod_os <> 0) " & SITUACAO & var_STATUS & " ORDER BY data_entrada, hora_entrada, OS.status;"
      
   ElseIf optData.Value = True Then
      If mskData.Text = "" Then Exit Sub
      sSQL = "SELECT cliente.*, OS.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' ELSE 'ABERTO' END AS var_status_os, OS.* " & _
         "FROM cliente INNER JOIN OS ON cliente.codigo = OS.cod_cliente WHERE (data_entrada = CONVERT(DATETIME, '" & Format$(mskData, ocDATA) & "', 103)) " & SITUACAO & var_STATUS & " ORDER BY data_entrada, hora_entrada, OS.status;"
      
   ElseIf optCliente.Value = True Then
      If CboCliente.Text = "" Then Exit Sub
      sSQL = "SELECT cliente.*, OS.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' ELSE 'ABERTO' END AS var_status_os, OS.* " & _
         "FROM cliente INNER JOIN OS ON cliente.codigo = OS.cod_cliente WHERE (cod_cliente = " & txtCodCliente.Text & ") " & SITUACAO & var_STATUS & " ORDER BY data_entrada, hora_entrada, OS.status;"
      
   ElseIf optMes.Value = True Then
      If cboMes.Text = "" Then Exit Sub
      sSQL = "SELECT cliente.*, OS.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' ELSE 'ABERTO' END AS var_status_os, OS.* " & _
         "FROM cliente INNER JOIN OS ON cliente.codigo = OS.cod_cliente WHERE (MONTH(data_entrada) = " & cboMes.ListIndex + 1 & ") AND (YEAR(data_entrada) = " & cboAno & ") " & SITUACAO & var_STATUS & _
         " ORDER BY data_entrada, hora_entrada, OS.status;"
      
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   
   FormatarGrid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   lblQuant.Caption = Format(totalRegistros, "00") & " ordem(ns) de serviço(s)"
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
   
   'For i = iAno To FirstYear Step -1
   '   cboAno.AddItem i
   'Next
   
   'iAno = iAno + 1
   'For i = iAno To LastYear
   '   cboAno.AddItem i
   'Next
End Sub

Private Sub cboAno_LostFocus()
   cmdExibirMes.SetFocus
End Sub

Private Sub cboCliente_Click()
   On Error GoTo TrataErro
   
   If CboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   If CboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = CboCliente.ItemData(CboCliente.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub CboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   CboCliente.Clear
   
   sSQL = "SELECT * FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      CboCliente.AddItem r("nome")
      CboCliente.ItemData(CboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo CboCliente
End Sub

Private Sub CboCliente_LostFocus()
   cboCliente_Click
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

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim aCor As ColorConstants
   
   With Grid
      .Rows = 1       'INICIA O GRID COM UMA LINHA
      .FixedCols = 0  'DETERMINA QUE NÃO HAJA COLUNA FIXA
      
      'Abaixo o cabeçalho é criado
      .FormatString = "^COD OS|^SITUAÇÃO|^STATUS|^CLIENTE|^AUTOMOVEL|^ENTRADA"
      .ColWidth(0) = 0
      .ColWidth(1) = 1300
      .ColWidth(2) = 1200
      .ColWidth(3) = 4000
      .ColWidth(4) = 2500
      .ColWidth(5) = 1500
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(3) = 1
            .ColAlignment(4) = 1
            .ColAlignment(5) = 1
            
            'A linha abaixo cria mais linha no grid
            .Rows = .Rows + 1
            
            'Preenche com os dados, e assim sucessivamente
            .TextMatrix(.Rows - 1, 0) = rTabela("cod_os")
            .TextMatrix(.Rows - 1, 1) = rTabela("var_status")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_status_os")
            .TextMatrix(.Rows - 1, 3) = rTabela("nome")
            .TextMatrix(.Rows - 1, 4) = rTabela("modelo")
            .TextMatrix(.Rows - 1, 5) = Format$(rTabela("data_entrada"), "dd/mm/yy") & " - " & Format$(rTabela("hora_entrada"), "hh:mm")
            
            rTabela.MoveNext
         Loop
      End If
      
      ' agora sim coloco a fução para mudar a cor da coluna e pronto
      'mudar a cor da fonte
      For i = 1 To .Rows - 1
         If UCase(Trim(.TextMatrix(i, 2))) = UCase("ABERTO") Then
            aCor = vbBlue
         Else
            aCor = vbRed
         End If
         
         .Col = 2 'a coluna do aberto ou fechado
         .Row = i
         .CellForeColor = aCor
      Next
      
      'mudar a cor da fonte
      For i = 1 To .Rows - 1
         If UCase(Trim(.TextMatrix(i, 1))) = UCase("À COMEÇAR") Then
            aCor = vbBlack
         ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("EM EXECUÇÃO") Then
            aCor = vbGreen
         ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("AGUARDANDO") Then
            aCor = vbYellow
         ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("TERMINADO") Then
            aCor = vbRed
         End If
         
         .Col = 1 'a coluna do aberto ou fechado
         .Row = i
         .CellForeColor = aCor
      Next
      
      .Redraw = True
   End With
End Sub

Private Sub cmdExibirCliente_Click()
   Montar_Grid
End Sub

Private Sub cmdExibirData_Click()
   Montar_Grid
End Sub

Private Sub cmdExibirMes_Click()
   Montar_Grid
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Montar_Grid
End Sub

Private Sub Form_Load()
   Set moCombo = New cComboHelper
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   Ordem_Servicos_Motores.SSTab1.Tab = 0
   Ordem_Servicos_Motores.frmSecundario.Enabled = True
   Ordem_Servicos_Motores.frmPrincipal.Enabled = True
   Ordem_Servicos_Motores.cmdGerarEntrada.Enabled = False
   Ordem_Servicos_Motores.cmdCancelarEntrada.Enabled = False
   Ordem_Servicos_Motores.cmdAlterar.Enabled = True
   Ordem_Servicos_Motores.cmdApagar.Enabled = True
   Ordem_Servicos_Motores.cmdNovo.Enabled = True
   Ordem_Servicos_Motores.txtCodOS.Text = ""
   Ordem_Servicos_Motores.txtCodOS.Text = (Grid.TextMatrix(Grid.Row, 0))
   Ordem_Servicos_Motores.Show
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub OptAbertos_Click()
   Montar_Grid
End Sub

Private Sub optCliente_Click()
   'desabilitar objetos
   lblMes.Enabled = False
   cboMes.Enabled = False
   cboAno.Enabled = False
   cmdExibirMes.Enabled = False
   
   lblData.Enabled = False
   mskData.Enabled = False
   cmdExibirData.Enabled = False
   
   lblCliente.Enabled = True
   CboCliente.Enabled = True
   cmdExibirCliente.Enabled = True
   CboCliente.SetFocus
End Sub

Private Sub optData_Click()
   'desabilitar objetos
   lblMes.Enabled = False
   cboMes.Enabled = False
   cboAno.Enabled = False
   cmdExibirMes.Enabled = False
   
   lblData.Enabled = True
   mskData.Enabled = True
   cmdExibirData.Enabled = True
   
   lblCliente.Enabled = False
   CboCliente.Enabled = False
   cmdExibirCliente.Enabled = False
   mskData.SetFocus
End Sub

Private Sub optFechados_Click()
   Montar_Grid
End Sub

Private Sub optFinTodos_Click()
   Montar_Grid
End Sub

Private Sub optMes_Click()
   'desabilitar objetos
   lblMes.Enabled = True
   cboMes.Enabled = True
   cboAno.Enabled = True
   cmdExibirMes.Enabled = True
   
   lblData.Enabled = False
   mskData.Enabled = False
   cmdExibirData.Enabled = False
   
   lblCliente.Enabled = False
   CboCliente.Enabled = False
   cmdExibirCliente.Enabled = False
   
   cboMes.SetFocus
End Sub

Private Sub optStatusAguardando_Click()
   Montar_Grid
End Sub

Private Sub optStatusComecar_Click()
   Montar_Grid
End Sub

Private Sub optStatusExecucao_Click()
   Montar_Grid
End Sub

Private Sub optStatusTerminado_Click()
   Montar_Grid
End Sub

Private Sub optStatusTodos_Click()
   Montar_Grid
End Sub

Private Sub optTodos_Click()
   Montar_Grid
   
   lblMes.Enabled = False
   cboMes.Enabled = False
   cboAno.Enabled = False
   cmdExibirMes.Enabled = False
   
   lblData.Enabled = False
   mskData.Enabled = False
   cmdExibirData.Enabled = False
   
   lblCliente.Enabled = False
   CboCliente.Enabled = False
   cmdExibirCliente.Enabled = False
End Sub
