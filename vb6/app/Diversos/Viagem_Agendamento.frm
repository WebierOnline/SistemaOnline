VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Viagem_Agendamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AGENDAMENTO DE VIAGENS"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   Icon            =   "Viagem_Agendamento.frx":0000
   LinkTopic       =   "Form73"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4575
      Left            =   60
      TabIndex        =   29
      Top             =   2760
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   8070
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11445
      TabIndex        =   22
      Top             =   60
      Width           =   11475
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9660
         TabIndex        =   25
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   8940
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Viagem_Agendamento.frx":23D2
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AGENDAMENTO"
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
         TabIndex        =   24
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mostrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   19
      Top             =   7620
      Width           =   3315
      Begin MSMask.MaskEdBox mskDataConsulta 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdAvancar 
         Caption         =   ">>"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "<<"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton Command12 
         Caption         =   "ok"
         Height          =   255
         Left            =   1620
         TabIndex        =   12
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame frmReserva 
      Caption         =   "AGENDAMENTO"
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
      Height          =   975
      Left            =   60
      TabIndex        =   15
      Top             =   1080
      Width           =   11475
      Begin VB.ComboBox cboPoltrona 
         Height          =   315
         Left            =   9180
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtCodCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         TabIndex        =   26
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboOrigem 
         Height          =   315
         Left            =   4980
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4815
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Left            =   9960
         TabIndex        =   5
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendario2 
         Height          =   315
         Left            =   11100
         TabIndex        =   32
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   3
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Viagem_Agendamento.frx":372D
         PICN            =   "Viagem_Agendamento.frx":3749
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Poltrona"
         Height          =   195
         Left            =   9180
         TabIndex        =   28
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   7080
         TabIndex        =   27
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Origem"
         Height          =   195
         Left            =   4980
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   9960
         TabIndex        =   17
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   480
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   555
      Left            =   9840
      TabIndex        =   10
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Viagem_Agendamento.frx":5B2B
      PICN            =   "Viagem_Agendamento.frx":5B47
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
      Height          =   555
      Left            =   60
      TabIndex        =   6
      Top             =   2100
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
      BTYPE           =   3
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Viagem_Agendamento.frx":5E61
      PICN            =   "Viagem_Agendamento.frx":5E7D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   555
      Left            =   1800
      TabIndex        =   7
      Top             =   2100
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
      BTYPE           =   3
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Viagem_Agendamento.frx":C747
      PICN            =   "Viagem_Agendamento.frx":C763
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   555
      Left            =   60
      TabIndex        =   8
      Top             =   2100
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "&Alterar"
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
      MICON           =   "Viagem_Agendamento.frx":13207
      PICN            =   "Viagem_Agendamento.frx":13223
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
      Height          =   555
      Left            =   1800
      TabIndex        =   9
      Top             =   2100
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
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
      MICON           =   "Viagem_Agendamento.frx":13AFD
      PICN            =   "Viagem_Agendamento.frx":13B19
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
      Height          =   555
      Left            =   8100
      TabIndex        =   0
      Top             =   2100
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "Viagem_Agendamento.frx":13E33
      PICN            =   "Viagem_Agendamento.frx":13E4F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCliente 
      Height          =   555
      Left            =   9840
      TabIndex        =   20
      Top             =   2100
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "&Cliente"
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
      MICON           =   "Viagem_Agendamento.frx":14B29
      PICN            =   "Viagem_Agendamento.frx":14B45
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
      TabIndex        =   21
      Top             =   8310
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16140
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21:34"
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
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   615
      Left            =   3420
      TabIndex        =   31
      Top             =   7620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
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
      MICON           =   "Viagem_Agendamento.frx":14E5F
      PICN            =   "Viagem_Agendamento.frx":14E7B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQuant 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   11310
      TabIndex        =   30
      Top             =   7380
      Width           =   225
   End
End
Attribute VB_Name = "Viagem_Agendamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private printSQL As String
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Private Function Atualizar_Dados() As Boolean
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE viagem_reserva SET " & _
      "cod_cliente = '" & txtCodCliente.Text & "', " & _
      "destino = '" & cboDestino.Text & "', " & _
      "origem = '" & cboOrigem.Text & "', " & _
      "data = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
      "poltrona = '" & cboPoltrona.Text & "' "
   
   'Condição para atualização
   sSQL = sSQL & "WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Limpar_Objetos()
txtCodigo.Text = ""
cboCliente.Text = ""
txtCodCliente.Text = ""
cboDestino.Text = ""
cboOrigem.Text = ""
cboPoltrona.Text = ""
mskData.Mask = ""
mskData.Text = ""
cmdNovo.Visible = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
frmReserva.Enabled = False
End Sub

Private Sub Mostrar_Grid()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

If Not IsDate(mskDataConsulta) Then Exit Sub
sSQL = "SELECT cliente.codigo, cliente.nome, cliente.endereco, cliente.bairro, cliente.ponto_de_referencia, cliente.celular, viagem_reserva.* FROM viagem_reserva INNER JOIN cliente ON cliente.codigo = viagem_reserva.cod_cliente WHERE (data = CONVERT(DATETIME, '" & Format(mskDataConsulta.Text, ocDATA) & "', 103)) ORDER BY bairro, endereco;"
   

Set r = dbData.OpenRecordset(sSQL, totalRegistros)

FormatarGrid r
If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL

'MOSTRAR A QUANTIDADE REGISTROS
lblQuant.Caption = Format(totalRegistros, "00")
End Sub

Private Sub Mostrar_Dados(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      cboOrigem.Text = rTabela("origem")
      cboDestino.Text = rTabela("destino")
      mskData.Text = Format(rTabela("data"), "dd/mm/yy")
      cboPoltrona.Text = rTabela("poltrona")
      txtCodCliente.Text = rTabela("cod_cliente")
   End If
End Sub
Private Sub cboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim itemAtual As String
   Dim codAtual As String
   
   itemAtual = cboCliente.Text
   codAtual = txtCodCliente.Text
   cboCliente.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCliente.AddItem r("nome")
      cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboCliente.Text = itemAtual
   txtCodCliente.Text = codAtual
   moCombo.AttachTo cboCliente
End Sub


Private Sub cboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub cboCliente_LostFocus()
  On Error GoTo TrataErro
   'If chkCodPedido.Value = Unchecked Then
   If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   'If chkCodPedido.Value = Unchecked Then
   If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
   
   txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
 '  If chkCodPedido.Value = Unchecked Then Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboDestino_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim itemAtual As String
   itemAtual = cboDestino.Text
   cboDestino.Clear
   
   sSQL = "SELECT destino FROM viagem_reserva GROUP BY destino;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDestino.AddItem ValidateNull(r("destino"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboDestino.Text = itemAtual
   moCombo.AttachTo cboDestino
End Sub

Private Sub cboDestino_KeyPress(KeyAscii As Integer)
KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub cboOrigem_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim itemAtual As String
   itemAtual = cboOrigem.Text
   cboOrigem.Clear
   
   sSQL = "SELECT origem FROM viagem_reserva GROUP BY origem;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboOrigem.AddItem ValidateNull(r("origem"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboOrigem.Text = itemAtual
   moCombo.AttachTo cboOrigem
End Sub


Private Sub cboOrigem_KeyPress(KeyAscii As Integer)
KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub cmdAlterar_Click()
   If txtCodigo.Text = "" Then Exit Sub
   
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Mostrar_Grid
End Sub

Private Sub cmdAvancar_Click()
Dim DataNova As Date
DataNova = Format(DateAdd("d", 1, mskDataConsulta), "dd/mm/yy")
mskDataConsulta.Text = Format(DataNova, "dd/mm/yy")
Mostrar_Grid
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
  
   mskData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos
cmdNovo.Visible = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = False
cmdExcluir.Visible = False
frmReserva.Enabled = False
End Sub

Private Sub cmdCliente_Click()
Clientes_Cadastro.Show
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboDestino.Text = "" Or txtCodCliente.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   Dim bRet As Boolean
   
   If txtCodigo.Text = "" Then Exit Sub
   
   'Solicita ao usuário confirmação da exclusão
   If ShowMsg("Excluir essa reserva?", vbInformation + vbYesNo) = vbNo Then Exit Sub

   'Faz a exclusão usando o comando DELETE do SQL
   sSQL = "DELETE FROM viagem_reserva WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      ShowMsg "Não foi possível excluir o registro.", vbCritical
      Exit Sub
   End If
   
   Limpar_Objetos
   Mostrar_Grid
End Sub

Private Sub cmdImprimir_Click()
Dim r As ADODB.Recordset
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Viagem_Reserva.Relatorio.Recordset = r

REL_Viagem_Reserva.rfQuant.Caption = lblQuant.Caption
REL_Viagem_Reserva.rfData.Caption = Format(mskDataConsulta.Text, "dd/mm/yy")

'REL_Viagem_Reserva.Relatorio.NomeImpressora = var_Impressora
REL_Viagem_Reserva.Relatorio.Ativar
Unload REL_Viagem_Reserva

Me.Show 1
End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
frmReserva.Enabled = True
mskData = Format(DateAdd("d", Val(1), Date), "dd/mm/yy")
cmdNovo.Visible = False
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdAlterar.Visible = False
cmdExcluir.Visible = False
cboCliente.SetFocus
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lNovoCod As Long
   
   If txtCodCliente.Text = "" Or cboDestino.Text = "" Or mskData.Text = "" Then
      ShowMsg "Formulário incompleto!", vbInformation
      cboCliente.SetFocus
      Exit Sub
   End If
   
   'ADICIONAR REGISTRO
   lNovoCod = AutoNumeracao
   
   'Faz a inserção de forma direta e verifica se houve algum erro
   If Not Inserir_Dados(lNovoCod) Then
      ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   Limpar_Objetos
   Mostrar_Grid
   Exit Sub
   
TrataErro:
   If Err.Number = 3022 Then
      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
      Exit Sub
   End If
End Sub


Private Sub cmdVoltar_Click()
Dim DataNova As Date
DataNova = Format(DateAdd("d", -1, mskDataConsulta), "dd/mm/yy")
mskDataConsulta.Text = Format(DataNova, "dd/mm/yy")
Mostrar_Grid
End Sub

Private Sub Command12_Click()
Mostrar_Grid
End Sub

Private Sub Form_Load()
frmReserva.Enabled = False
mskDataConsulta.Text = Format(DateAdd("d", Val(1), Date), "dd/mm/yy")
Mostrar_Grid
Set moCombo = New cComboHelper
End Sub
Private Sub FormatarGrid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 3800
      .ColWidth(4) = 2000
      .ColWidth(5) = 2000
      .ColWidth(6) = 1700
      .ColWidth(7) = 700
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "COD_CLIENTE"
      .TextMatrix(0, 3) = "CLIENTE"
      .TextMatrix(0, 4) = "BAIRRO"
      .TextMatrix(0, 5) = "ORIGEM"
      .TextMatrix(0, 6) = "DESTINO"
      .TextMatrix(0, 7) = "POLT."
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("cod_cliente")
            .TextMatrix(.Rows - 1, 3) = rTabela("nome")
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("bairro"))
            .TextMatrix(.Rows - 1, 5) = rTabela("origem")
            .TextMatrix(.Rows - 1, 6) = rTabela("destino")
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("poltrona"), "00")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 4
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   'lblValor.Caption = Format(SomaGrid(GridSuprimentos, 6), ocMONEY)
End Sub
Private Function Inserir_Dados(ByVal Codigo As Long) As Boolean
   'A inclusão deve ser feita utilizando o comando INSERT INTO do sql
   'e não mais usando o método .AddNew do Recordset
   
   Dim sSQL As String
   
   'Comando de inclusão
   sSQL = "INSERT INTO viagem_reserva (codigo, cod_cliente, origem, destino, data, poltrona) VALUES (" & _
      Codigo & ", " & txtCodCliente.Text & ", '" & cboOrigem.Text & "', '" & cboDestino.Text & "', CONVERT(DATETIME, '" & _
      Format$(mskData.Text, ocDATA) & "', 103), '" & cboPoltrona.Text & "');"
   
   'Retorna o resultado da inclusão
   Inserir_Dados = dbData.Execute(sSQL)
End Function
Private Function AutoNumeracao() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_reserva FROM viagem_reserva;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then lRet = r("cod_reserva") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao = lRet
End Function

Function Maiuscula(KeyAscii As Integer)
   If KeyAscii > 96 And KeyAscii < 123 Then
      KeyAscii = KeyAscii - 32
   End If
   Maiuscula = KeyAscii
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
   'TBcompromissos.Close
End Sub

Private Sub Grid_DblClick()
frmReserva.Enabled = True
cmdNovo.Visible = True
cmdSalvar.Visible = False
cmdCancelar.Visible = False
cmdAlterar.Visible = True
cmdExcluir.Visible = True
txtCodigo.Text = ""
txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub


Private Sub mskData_KeyPress(KeyAscii As Integer)
mskData.Mask = "##/##/##"
End Sub


Private Sub mskDataConsulta_Validate(Cancel As Boolean)
   If mskDataConsulta.Text = "__/__/__" Then
      mskDataConsulta.SetFocus
      Exit Sub
   End If
   
   If Not IsDate(mskDataConsulta) Then
      ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
      mskDataConsulta.SetFocus
      mskDataConsulta.SelStart = 0
      mskDataConsulta.SelLength = Len(mskDataConsulta)
   Exit Sub
   End If
End Sub


Private Sub txtCodCliente_Change()
If txtCodCliente.Text = "" Then Exit Sub
Dim r As ADODB.Recordset
If cmdAlterar.Visible = True Then
      sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then cboCliente.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtCodigo_Change()
   If txtCodigo.Text = "" Then Exit Sub
   
   If cmdAlterar.Visible = True Then
      sSQL = "SELECT * FROM viagem_reserva WHERE (codigo = " & txtCodigo.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If Not r.BOF Then Mostrar_Dados r
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub


