VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Carne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CARNÊ"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "Carne.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   4545
      TabIndex        =   9
      Top             =   60
      Width           =   4575
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Carne.frx":23D2
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CARNÊ"
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
         Left            =   1140
         TabIndex        =   10
         Top             =   120
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   60
      ScaleHeight     =   1995
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   780
      Width           =   4575
      Begin VB.ComboBox cboParcela 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   300
         Width           =   4395
      End
      Begin VB.ComboBox cboPedido 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   915
         Width           =   2175
      End
      Begin VB.Label lblParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
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
         Left            =   60
         TabIndex        =   8
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Left            =   75
         TabIndex        =   7
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedido | Data"
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
         Left            =   75
         TabIndex        =   6
         Top             =   675
         Width           =   1170
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2940
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Sair"
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
      MICON           =   "Carne.frx":8742
      PICN            =   "Carne.frx":875E
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
      Left            =   60
      TabIndex        =   3
      Top             =   2940
      Width           =   1635
      _ExtentX        =   2884
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
      MICON           =   "Carne.frx":8A78
      PICN            =   "Carne.frx":8A94
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
      TabIndex        =   11
      Top             =   3675
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3969
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:15"
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
Attribute VB_Name = "Carne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Private Sub cboCliente_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lCodigo As Long
   
   If cboCliente.Text = "" Then Exit Sub
   
   lCodigo = 0
   sSQL = "SELECT * FROM cliente WHERE (nome = '" & cboCliente.Text & "');"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lCodigo = r("codigo")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   sSQL = "SELECT * FROM pedidos WHERE (cod_cliente = " & lCodigo & ") ORDER BY data_compra DESC;"
   Set r = dbData.OpenRecordset(sSQL)
   
   cboPedido.Clear
   
   If r.BOF Then
      cboPedido.AddItem "NENHUM PEDIDO"
   Else
      Do While Not r.EOF
         cboPedido.AddItem Format(r("cod_pedido"), "00000") & " -> " & r("data_compra")
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
   
   cboPedido.ListIndex = 0
End Sub

Private Sub CboCliente_GotFocus()
   moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cboCliente_Click
End Sub

Private Sub CboCliente_LostFocus()
   cboCliente_Click
End Sub

Private Sub cboPedido_Change()
   cboPedido_Click
End Sub

Private Sub cboPedido_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim Pedido As Long

cboParcela.Clear

If cboPedido.Text = "NENHUM PEDIDO" Then
   Pedido = 0
Else
   Pedido = Mid(cboPedido.Text, 1, InStr(1, cboPedido.Text, "->") - 1)
End If
 
sSQL = "SELECT * FROM parcelas INNER JOin pedidos ON parcelas.cod_pedido = pedidos.cod_pedido WHERE (pedidos.pagamento = 'Promissoria') AND (pedidos.cod_pedido = " & Pedido & ");"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   cboParcela.AddItem "SEM PROMISSÓRIA"
Else
   cboParcela.AddItem "TODAS"
   Do While Not r.EOF
      cboParcela.AddItem r("numero")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If

cboParcela.ListIndex = 0
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim Pedido As Integer

If cboPedido.Text = "" Or cboCliente.Text = "" Or cboParcela.Text = "" Then Exit Sub

If cboPedido.Text = "NENHUM PEDIDO" Then
   ShowMsg "ESSE PEDIDO NÃO PODE SER IMPRESSO", vbExclamation
   Exit Sub
End If

If cboParcela.Text = "SEM PROMISSÓRIA" Then
   ShowMsg "ESSE PEDIDO NÃO PODE SER IMPRESSO", vbExclamation
   Exit Sub
End If

Me.Hide

'Principal_Impressao.Hide
Pedido = Mid(cboPedido.Text, 1, InStr(1, cboPedido.Text, "->") - 1)

With REL_Carne_ContinuoG
   '.loadPromissoria Pedido, 1
   .loadPromissoria Pedido, IIf(cboParcela.ListIndex = 0, 0, cboParcela.Text)
End With

'   Unload Imp_Detalhes_Pedido
Unload Me

Me.Show 1
End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "select * from Cliente order by Nome"
Set r = dbData.OpenRecordset(sSQL)

cboCliente.Clear
Do While Not r.EOF
   cboCliente.AddItem r("nome")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'cboCliente.ListIndex = 0
Set moCombo = New cComboHelper
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub
