VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNFSe_OS
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nova NFS-e - OS finalizadas sem nota"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCliente
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdFiltrar
      Caption         =   "&Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   90
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdOS
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdGerar
      Caption         =   "&Gerar NFS-e"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdFechar
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   8385
      TabIndex        =   7
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label lblCli
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   615
   End
   Begin VB.Label lblInfo
      Caption         =   "Clique na primeira coluna (ou dê dois cliques) para marcar as OS. Todas as OS de uma NFS-e precisam ser do mesmo cliente. Só os SERVIÇOS vão na NFS-e (as peças vão na NF-e)."
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   9960
   End
   Begin VB.Label lblTotal
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
      Left            =   120
      TabIndex        =   5
      Top             =   6180
      Width           =   6375
   End
End
Attribute VB_Name = "frmNFSe_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Escolha das OS finalizadas sem NFS-e (view NFSe_OSDisponiveis) e geração da nota pelo NFSeEmissor (acao=gerar).
Public IdGerado As Long

Private Const C_X = 0, C_OS = 1, C_DATA = 2, C_CLI = 3, C_SERV = 4, C_DESC = 5, C_LIQ = 6, C_CODCLI = 7

Private Sub Form_Load()
IdGerado = 0
Carregar
End Sub

Private Sub Carregar()
Dim r As ADODB.Recordset
Dim sSQL As String
Dim n As Long

With grdOS
   .Rows = 1
   .Cols = 8
   .TextMatrix(0, C_X) = "X"
   .TextMatrix(0, C_OS) = "OS"
   .TextMatrix(0, C_DATA) = "Finalizada"
   .TextMatrix(0, C_CLI) = "Cliente"
   .TextMatrix(0, C_SERV) = "Serviços"
   .TextMatrix(0, C_DESC) = "Desconto"
   .TextMatrix(0, C_LIQ) = "Líquido"
   .ColWidth(C_X) = 300
   .ColWidth(C_OS) = 900
   .ColWidth(C_DATA) = 1100
   .ColWidth(C_CLI) = 4100
   .ColWidth(C_SERV) = 1150
   .ColWidth(C_DESC) = 1050
   .ColWidth(C_LIQ) = 1150
   .ColWidth(C_CODCLI) = 0
   .ColAlignment(C_X) = flexAlignCenterCenter
   .ColAlignment(C_SERV) = flexAlignRightCenter
   .ColAlignment(C_DESC) = flexAlignRightCenter
   .ColAlignment(C_LIQ) = flexAlignRightCenter
End With

sSQL = "SELECT TOP 500 COD_OS, COD_CLIENTE, Cliente, DATA_TERMINO, ValorServicos, DescontoServicos FROM NFSe_OSDisponiveis"
If Trim$(txtCliente.Text) <> "" Then sSQL = sSQL & " WHERE (Cliente LIKE " & NFSe_Q("%" & Trim$(txtCliente.Text) & "%") & ")"
sSQL = sSQL & " ORDER BY DATA_TERMINO DESC, COD_OS DESC"
Set r = dbData.OpenRecordset(sSQL)
With grdOS
   .Redraw = False
   Do While Not r.EOF
      n = .Rows
      .Rows = n + 1
      .TextMatrix(n, C_OS) = r("COD_OS")
      .TextMatrix(n, C_DATA) = IIf(IsNull(r("DATA_TERMINO")), "", Format$(r("DATA_TERMINO"), "dd/mm/yyyy"))
      .TextMatrix(n, C_CLI) = ValidateNull(r("Cliente"))
      .TextMatrix(n, C_SERV) = Format$(ValidateNull(r("ValorServicos")), ocMONEY)
      .TextMatrix(n, C_DESC) = Format$(ValidateNull(r("DescontoServicos")), ocMONEY)
      .TextMatrix(n, C_LIQ) = Format$(CDbl(ValidateNull(r("ValorServicos"))) - CDbl(ValidateNull(r("DescontoServicos"))), ocMONEY)
      .TextMatrix(n, C_CODCLI) = ValidateNull(r("COD_CLIENTE"))
      r.MoveNext
   Loop
   .Redraw = True
End With
r.Close
Set r = Nothing
AtualizarTotal
End Sub

Private Sub Marcar(ByVal nLinha As Long)
Dim i As Long
If nLinha < 1 Then Exit Sub
If grdOS.TextMatrix(nLinha, C_X) = "X" Then
   grdOS.TextMatrix(nLinha, C_X) = ""
Else
   For i = 1 To grdOS.Rows - 1
      If grdOS.TextMatrix(i, C_X) = "X" And grdOS.TextMatrix(i, C_CODCLI) <> grdOS.TextMatrix(nLinha, C_CODCLI) Then
         MsgBox "Essa OS é de outro cliente." & vbCrLf & "Uma NFS-e só pode ter OS do mesmo cliente.", vbExclamation, "NFS-e"
         Exit Sub
      End If
   Next
   grdOS.TextMatrix(nLinha, C_X) = "X"
End If
AtualizarTotal
End Sub

Private Sub AtualizarTotal()
Dim i As Long, nQtd As Long
Dim dTotal As Double
For i = 1 To grdOS.Rows - 1
   If grdOS.TextMatrix(i, C_X) = "X" Then
      nQtd = nQtd + 1
      dTotal = dTotal + CDbl(grdOS.TextMatrix(i, C_LIQ))
   End If
Next
lblTotal.Caption = IIf(nQtd = 0, "Nenhuma OS marcada", nQtd & " OS marcada(s) - serviços: R$ " & Format$(dTotal, ocMONEY))
cmdGerar.Enabled = (nQtd > 0)
End Sub

Private Sub grdOS_Click()
If grdOS.MouseCol = C_X Then Marcar grdOS.MouseRow
End Sub

Private Sub grdOS_DblClick()
If grdOS.MouseCol <> C_X Then Marcar grdOS.Row
End Sub

Private Sub grdOS_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then KeyAscii = 0: Marcar grdOS.Row
End Sub

Private Sub cmdFiltrar_Click()
Carregar
End Sub

Private Sub cmdGerar_Click()
Dim oRes As Object
Dim sStatus As String, sLista As String
Dim i As Long
For i = 1 To grdOS.Rows - 1
   If grdOS.TextMatrix(i, C_X) = "X" Then sLista = sLista & IIf(sLista = "", "", ",") & grdOS.TextMatrix(i, C_OS)
Next
If sLista = "" Then Exit Sub
sStatus = NFSe_Executar("acao=gerar" & vbCrLf & "os=" & sLista, oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível gerar a NFS-e:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
   Exit Sub
End If
IdGerado = Val(NFSe_Valor(oRes, "idnfse"))
Me.Hide
End Sub

Private Sub cmdFechar_Click()
IdGerado = 0
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
   Cancel = True
   cmdFechar_Click
End If
End Sub
