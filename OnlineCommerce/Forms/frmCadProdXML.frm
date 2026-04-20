VERSION 5.00
Begin VB.Form frmCadProdXML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO DE PRODUTOS - VAREJO"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInfo 
      Caption         =   "Produto da Nota Fiscal de Compra:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8160
      Begin VB.Label lblNomeXML 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   900
         TabIndex        =   1
         Top             =   300
         Width           =   7020
      End
      Begin VB.Label lblUnid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1260
         TabIndex        =   2
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label lblNCM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblCEST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label lblValor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1260
         TabIndex        =   5
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label lblICMSCST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblPISCST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5040
         TabIndex        =   7
         Top             =   900
         Width           =   765
      End
      Begin VB.Label lblCOFINSCST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   8
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid. Medida:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM:"
         Height          =   225
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEST:"
         Height          =   225
         Left            =   4920
         TabIndex        =   12
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Unit.:"
         Height          =   195
         Left            =   420
         TabIndex        =   13
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS CST:"
         Height          =   225
         Left            =   2640
         TabIndex        =   14
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIS CST:"
         Height          =   225
         Left            =   4320
         TabIndex        =   15
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COFINS CST:"
         Height          =   225
         Left            =   5880
         TabIndex        =   16
         Top             =   900
         Width           =   810
      End
   End
   Begin VB.Frame fraCad 
      Caption         =   "Produto que você está cadastando (Varejo):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   60
      TabIndex        =   17
      Top             =   1500
      Width           =   8160
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   390
         Width           =   6840
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   780
         Width           =   1200
      End
      Begin VB.TextBox txtEAN 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1440
         MaxLength       =   14
         TabIndex        =   21
         Top             =   1410
         Width           =   2400
      End
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COD. BARRA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblEANHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Códifo de barra da embalagem UNITARIA (ex: 7892840808051)"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   1140
         Width           =   4530
      End
      Begin VB.Label lblHint1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leia o codigo de barras da embalagem individual (nao da caixa do fornecedor)."
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   1740
         Width           =   5580
      End
      Begin VB.Label lblAviso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Sem o COD. BARRA o produto não será salvo."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   24
         Top             =   1980
         Width           =   3435
      End
      Begin VB.Label lblRegime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   7665
         TabIndex        =   25
         Top             =   2280
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Default         =   -1  'True
      Height          =   435
      Left            =   5280
      TabIndex        =   26
      Top             =   4080
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6780
      TabIndex        =   27
      Top             =   4080
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
   Cancelado = True
   ResDescricao = ""
   ResEAN = ""
End Sub

'==============================================================
'Chamado pelo formulario pai apos definir as variaveis Pub*
Public Sub InicializarUI()
   lblNomeXML.Caption = PubNome
   lblUnid.Caption = PubUnidade
   lblNCM.Caption = PubNCM
   lblCEST.Caption = PubCEST
   lblValor.Caption = FormatNumber(PubValorUnit, 2)
   lblICMSCST.Caption = PubICMSCST
   lblPISCST.Caption = PubPISCST
   lblCOFINSCST.Caption = PubCOFINSCST

   Dim sRegNome As String
   If PubRegime = 1 Or PubRegime = 2 Then
      sRegNome = "Simples Nacional (ICMS 102 / PIS 07 / COFINS 07)"
   Else
      sRegNome = "Regime Normal (tributacao conforme NF-e de entrada)"
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
   Dim sEANBusca As String
   sEANBusca = Replace(Trim(txtEAN.Text), "'", "''")
   Dim lQtdExist As Long
   lQtdExist = SQLExecutaRetorno("SELECT COUNT(*) r FROM Produtos WHERE LTRIM(RTRIM(ISNULL(COD_BARRA,''))) = '" & sEANBusca & "' OR LTRIM(RTRIM(ISNULL(EAN,''))) = '" & sEANBusca & "'", "r", 0)
   If lQtdExist > 0 Then
      Dim lCodExist As Long
      lCodExist = SQLExecutaRetorno("SELECT TOP 1 Codigo r FROM Produtos WHERE LTRIM(RTRIM(ISNULL(COD_BARRA,''))) = '" & sEANBusca & "' OR LTRIM(RTRIM(ISNULL(EAN,''))) = '" & sEANBusca & "'", "r", 0)
      Dim sDescExist As String
      sDescExist = SQLExecutaRetorno("SELECT ISNULL(DESCRICAO,'') r FROM Produtos WHERE Codigo = " & lCodExist, "r", "")
      MsgBox "Ja existe um produto cadastrado com este codigo de barras:" & vbCrLf & vbCrLf & _
             "EAN: " & Trim(txtEAN.Text) & vbCrLf & _
             "Descricao: " & sDescExist & vbCrLf & vbCrLf & _
             "Verifique se o produto ja esta cadastrado antes de criar um novo.", _
             vbExclamation, "Codigo de Barras Duplicado"
      txtEAN.SetFocus
      Exit Sub
   End If
   Cancelado = False
   ResDescricao = Trim(txtDescricao.Text)
   ResEAN = Trim(txtEAN.Text)
   ResUnidade = cboUnidade.Text
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
