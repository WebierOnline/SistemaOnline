VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBuscarPlaca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Veículo por Placa"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid lstVeiculos 
      Height          =   3540
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   6244
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtPlacaF 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   150
      Width           =   1800
   End
   Begin VB.Frame fraCarregando 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3500
      TabIndex        =   3
      Top             =   1900
      Visible         =   0   'False
      Width           =   2800
      Begin VB.Label lblCarregando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   4
         Top             =   90
         Width           =   2680
      End
   End
   Begin VB.Timer tmrCarregando 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2880
      Top             =   3900
   End
   Begin ChamaleonBtn.chameleonButton cmdHistorico 
      Height          =   315
      Left            =   7260
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Histórico"
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
      MICON           =   "frmBuscarPlaca.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdUsarEsse 
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "USAR ESSE"
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
      MICON           =   "frmBuscarPlaca.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFiltrar 
      Height          =   315
      Left            =   2460
      TabIndex        =   7
      Top             =   180
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Buscar"
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
      MICON           =   "frmBuscarPlaca.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   315
      Left            =   8520
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Fechar"
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
      MICON           =   "frmBuscarPlaca.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPlacaF 
      Caption         =   "Placa:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmBuscarPlaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sPlacaSel As String
Public lCodClienteSel As Long
Public sNomeClienteSel As String
Public sCelularClienteSel As String
Public sModeloSel As String
Public sAnoSel As String
Public sKmSel As String
Public sCorSel As String
Public sChassiSel As String
Public lCodOSSelecionado As Long

Private iDots As Integer

Private Sub Form_Load()
    sPlacaSel = ""
    iDots = 0
    fraCarregando.Visible = False
    ConfigurarGrid
End Sub

Private Sub Form_Activate()
    txtPlacaF.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        sPlacaSel = ""
        Me.Hide
    End If
End Sub

Private Sub ConfigurarGrid()
    With lstVeiculos
        .Cols = 10
        .Rows = 1
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 3400
        .ColWidth(3) = 1300
        .ColWidth(4) = 700
        .ColWidth(5) = 900
        .ColWidth(6) = 700
        .ColWidth(7) = 1000
        .ColWidth(8) = 1800
        .ColWidth(9) = 0
        .Row = 0: .Col = 2: .Text = "CLIENTE"
        .Row = 0: .Col = 3: .Text = "MODELO"
        .Row = 0: .Col = 4: .Text = "ANO"
        .Row = 0: .Col = 5: .Text = "PLACA": .CellBackColor = &HE0E0E0
        .Row = 0: .Col = 6: .Text = "KM"
        .Row = 0: .Col = 7: .Text = "COR"
        .Row = 0: .Col = 8: .Text = "CHASSI"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
End Sub

Private Sub CarregarGrid()
    Dim rVei As ADODB.Recordset
    Dim sql As String
    Dim n As Integer

    If Trim(txtPlacaF.Text) = "" Then
        MsgBox "Digite a placa para consultar!", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    fraCarregando.Visible = True
    fraCarregando.ZOrder 0
    tmrCarregando.Enabled = True
    Me.Refresh
    DoEvents

    lstVeiculos.Rows = 1

    sql = "SELECT DISTINCT cliente.codigo AS cod_cliente, cliente.nome, cliente.celular, " & _
          "OS_Equipamento_Auto.modelo, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.placa, " & _
          "OS_Equipamento_Auto.km, OS_Equipamento_Auto.cor, OS_Equipamento_Auto.chassi, OS.COD_OS AS cod_os " & _
          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE " & _
          "INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _
          "WHERE (OS_Equipamento_Auto.placa LIKE '%" & Replace(Trim(txtPlacaF.Text), "'", "''") & "%') " & _
          "ORDER BY cliente.nome"

    RsOpen rVei, sql
    n = 1
    Do While Not rVei.EOF
        With lstVeiculos
            .Rows = n + 1
            .Row = n: .Col = 0: .Text = ValidateNull(rVei("cod_cliente"))
            .Row = n: .Col = 1: .Text = ValidateNull(rVei("celular"))
            .Row = n: .Col = 2: .Text = ValidateNull(rVei("nome"))
            .Row = n: .Col = 3: .Text = ValidateNull(rVei("modelo"))
            .Row = n: .Col = 4: .Text = ValidateNull(rVei("ano"))
            .Row = n: .Col = 5: .Text = ValidateNull(rVei("placa"))
            .Row = n: .Col = 6: .Text = ValidateNull(rVei("km"))
            .Row = n: .Col = 7: .Text = ValidateNull(rVei("cor"))
            .Row = n: .Col = 8: .Text = ValidateNull(rVei("chassi"))
            .Row = n: .Col = 9: .Text = ValidateNull(rVei("cod_os"))
        End With
        n = n + 1
        rVei.MoveNext
    Loop
    If rVei.State <> 0 Then rVei.Close

    tmrCarregando.Enabled = False
    fraCarregando.Visible = False

    If lstVeiculos.Rows = 1 Then
        MsgBox "Nenhum veículo encontrado para essa placa!", vbInformation, "Aviso do Sistema"
    End If
End Sub

Private Sub txtPlacaF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        CarregarGrid
    End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdFiltrar_Click()
    CarregarGrid
End Sub

Private Sub lstVeiculos_DblClick()
    cmdUsarEsse_Click
End Sub

Private Sub cmdUsarEsse_Click()
    If lstVeiculos.Row < 1 Then
        MsgBox "Selecione um Veículo.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If
    lCodClienteSel = Val(lstVeiculos.TextMatrix(lstVeiculos.Row, 0))
    sCelularClienteSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 1)
    sNomeClienteSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 2)
    sModeloSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 3)
    sAnoSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 4)
    sPlacaSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 5)
    sKmSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 6)
    sCorSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 7)
    sChassiSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 8)
    lCodOSSelecionado = Val(lstVeiculos.TextMatrix(lstVeiculos.Row, 9))
    Me.Hide
End Sub

Private Sub cmdHistorico_Click()
    If lstVeiculos.Row < 1 Then
        MsgBox "Selecione um veículo.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    Dim sPlaca As String
    sPlaca = Trim(lstVeiculos.TextMatrix(lstVeiculos.Row, 5))
    If sPlaca = "" Then
        MsgBox "Veículo sem placa cadastrada.", vbInformation, "Aviso do Sistema"
        Exit Sub
    End If

    OS_Consulta.sPlacaBusca = sPlaca
    OS_Consulta.Show vbModal
    Unload OS_Consulta
End Sub

Private Sub cmdFechar_Click()
    sPlacaSel = ""
    Me.Hide
End Sub

Private Sub tmrCarregando_Timer()
    iDots = (iDots Mod 3) + 1
    lblCarregando.Caption = "Carregando" & String(iDots, ".")
End Sub
