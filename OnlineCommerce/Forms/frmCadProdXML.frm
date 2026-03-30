VERSION 5.00
Begin VB.Form frmCadProdXML
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastrar Produto - Varejo"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInfo
      Caption         =   "Produto conforme NF-e"
      Height          =   2040
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8160
      Begin VB.Label lblNomeXML
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   6240
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblUnid
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   1200
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblNCM
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   1680
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCEST
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   720
         Width           =   2280
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblValor
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1440
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblICMSCST
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   720
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPISCST
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   1080
         Width           =   720
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCOFINSCST
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6960
         TabIndex        =   8
         Top             =   1080
         Width           =   960
         BeginProperty Font
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome XML:"
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label2
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade:"
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label3
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM:"
         Height          =   225
         Left            =   2640
         TabIndex        =   11
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label4
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEST:"
         Height          =   225
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label5
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vl.Unit.:"
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label6
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS CST:"
         Height          =   225
         Left            =   2760
         TabIndex        =   14
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label7
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIS CST:"
         Height          =   225
         Left            =   4440
         TabIndex        =   15
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label8
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COFINS CST:"
         Height          =   225
         Left            =   6120
         TabIndex        =   16
         Top             =   1080
         Width           =   810
      End
   End
   Begin VB.Frame fraCad
      Caption         =   "Dados para Cadastro (Varejo)"
      Height          =   2820
      Left            =   120
      TabIndex        =   17
      Top             =   2220
      Width           =   8160
      Begin VB.Label Label9
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descricao:"
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   420
         Width           =   870
      End
      Begin VB.TextBox txtDescricao
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   390
         Width           =   6600
      End
      Begin VB.Label Label11
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade:"
         Height          =   225
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   750
      End
      Begin VB.ComboBox cboUnidade
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   810
         Width           =   1200
      End
      Begin VB.Label Label10
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COD. BARRA *:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   1260
         Width           =   1170
      End
      Begin VB.TextBox txtEAN
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   21
         Top             =   1230
         Width           =   2400
      End
      Begin VB.Label lblEANHint
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN da embalagem UNITARIA (ex: 7892840808051)"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   4080
         TabIndex        =   22
         Top             =   1275
         Width           =   3840
      End
      Begin VB.Label lblHint1
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leia o codigo de barras da embalagem individual (nao da caixa do fornecedor)."
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   7680
      End
      Begin VB.Label lblAviso
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Sem o COD. BARRA o produto nao sera salvo."
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   24
         Top             =   1980
         Width           =   7680
      End
      Begin VB.Label lblRegime
         BackStyle       =   0  'Transparent
         Caption         =   ""
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   2280
         Width           =   7680
      End
   End
   Begin VB.CommandButton cmdSalvar
      Caption         =   "&Salvar"
      Default         =   -1  'True
      Height          =   435
      Left            =   5400
      TabIndex        =   26
      Top             =   5220
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancelar
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6960
      TabIndex        =   27
      Top             =   5220
      Width           =   1440
   End
End
Attribute VB_Name = "frmCadProdXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' frmCadProdXML
' Dialog para cadastrar produto no varejo a partir de dados da NF-e.
' O COD_BARRA (EAN unitario) e obrigatorio antes de salvar.
'==============================================================
Option Explicit

'--- Variaveis de entrada (definidas pelo chamador antes de InicializarUI) ---
Public PubNome       As String
Public PubUnidade    As String
Public PubNCM        As String
Public PubCEST       As String
Public PubICMSCST    As String
Public PubPISCST     As String
Public PubCOFINSCST  As String
Public PubValorUnit  As Double
Public PubRegime     As Integer   '1/2=Simples Nacional, 3=Regime Normal

'--- Resultados (lidos pelo chamador apos Show retornar) ---
Public Cancelado     As Boolean
Public ResDescricao  As String
Public ResEAN        As String
Public ResUnidade    As String

'==============================================================
Private Sub Form_Load()
   Cancelado    = True
   ResDescricao = ""
   ResEAN       = ""
End Sub

'==============================================================
'Chamado pelo formulario pai apos definir as variaveis Pub*
Public Sub InicializarUI()
   lblNomeXML.Caption   = PubNome
   lblUnid.Caption      = PubUnidade
   lblNCM.Caption       = PubNCM
   lblCEST.Caption      = PubCEST
   lblValor.Caption     = "R$ " & Format(PubValorUnit, "##,##0.0000")
   lblICMSCST.Caption   = PubICMSCST
   lblPISCST.Caption    = PubPISCST
   lblCOFINSCST.Caption = PubCOFINSCST

   Dim sRegNome As String
   If PubRegime = 1 Or PubRegime = 2 Then
      sRegNome = "Simples Nacional  (ICMS 102 / PIS 07 / COFINS 07)"
   Else
      sRegNome = "Regime Normal  (tributacao conforme NF-e de entrada)"
   End If
   lblRegime.Caption = "Tributacao aplicada: " & sRegNome

   'Popula combobox de unidades
   Dim aUnits() As String
   Dim i As Integer
   aUnits = Split("UN,CX,KG,CT,PO,SC,PA,EX,BJ,DZ,PC,DI,FD,PT,M2,M3", ",")
   cboUnidade.Clear
   For i = 0 To UBound(aUnits)
      cboUnidade.AddItem aUnits(i)
   Next i
   'Seleciona unidade padrao (PubUnidade ou "UN")
   Dim sSelUnit As String
   sSelUnit = PubUnidade
   If sSelUnit = "" Then sSelUnit = "UN"
   cboUnidade.ListIndex = 0  'default UN
   For i = 0 To cboUnidade.ListCount - 1
      If cboUnidade.List(i) = sSelUnit Then
         cboUnidade.ListIndex = i
         Exit For
      End If
   Next i

   txtDescricao.Text = PubNome
   txtEAN.Text = ""
End Sub

'==============================================================
Private Sub Form_Activate()
   txtEAN.SetFocus
End Sub

'==============================================================
Private Sub cmdSalvar_Click()
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "A descricao e obrigatoria.", vbExclamation, "Campo Obrigatorio"
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Trim(txtEAN.Text) = "" Then
      MsgBox "O COD. BARRA (EAN unitario) e obrigatorio." & vbCrLf & _
             "Digite ou leia o codigo de barras da embalagem individual do produto.", _
             vbExclamation, "Campo Obrigatorio"
      txtEAN.SetFocus
      Exit Sub
   End If
   Cancelado    = False
   ResDescricao = Trim(txtDescricao.Text)
   ResEAN       = Trim(txtEAN.Text)
   ResUnidade   = cboUnidade.Text
   Me.Hide
End Sub

'==============================================================
Private Sub cmdCancelar_Click()
   Cancelado = True
   Me.Hide
End Sub

'==============================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      Cancelado = True
      Cancel = 1
      Me.Hide
   End If
End Sub
