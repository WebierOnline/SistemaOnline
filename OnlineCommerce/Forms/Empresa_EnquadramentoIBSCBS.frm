VERSION 5.00
Begin VB.Form Empresa_EnquadramentoIBSCBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enquadramento de IBS/CBS (NT 2026.005)"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblCRT 
      BackStyle       =   0  'Transparent
      Caption         =   "CRT"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   6720
   End
   Begin VB.Label lblAviso 
      BackStyle       =   0  'Transparent
      Caption         =   "Estes dados serão usados na emissão da NF-e/NFC-e conforme a NT 2026.005 (Simples Nacional). A emissão ainda não usa estes campos: salvar aqui não altera as notas por enquanto."
      ForeColor       =   &H00606060&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6720
   End
   Begin VB.Frame fraApuracao 
      Caption         =   "Apuração de IBS/CBS"
      Height          =   1860
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6720
      Begin VB.CheckBox chkRegular 
         Caption         =   "Optei pelo Regime Regular de IBS/CBS (recolho fora do DAS)"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   6240
      End
      Begin VB.Label lblPendente 
         BackStyle       =   0  'Transparent
         Caption         =   "Opção pendente pelo Simples:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2200
      End
      Begin VB.ComboBox cboPendente 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   795
         Width           =   3960
      End
      Begin VB.Label lblResumo 
         BackStyle       =   0  'Transparent
         Caption         =   ""
         ForeColor       =   &H00804000&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1260
         Width           =   6240
      End
   End
   Begin VB.Frame fraAtiv 
      Caption         =   "Atividade padrão do Simples (cAtivSN)"
      Height          =   1260
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   6720
      Begin VB.ComboBox cboAtiv 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   6240
      End
      Begin VB.Label lblAtivInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Usada como padrão em produtos novos e nos produtos sem atividade definida. Cada produto pode ter a sua."
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   6240
      End
   End
   Begin VB.CommandButton cmdRBT12 
      Caption         =   "RBT12 e alíquotas..."
      Height          =   435
      Left            =   240
      TabIndex        =   12
      Top             =   4620
      Width           =   2400
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Default         =   -1  'True
      Height          =   435
      Left            =   3720
      TabIndex        =   10
      Top             =   4620
      Width           =   1575
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   435
      Left            =   5400
      TabIndex        =   11
      Top             =   4620
      Width           =   1560
   End
End
Attribute VB_Name = "Empresa_EnquadramentoIBSCBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Enquadramento do emitente para IBS/CBS (NT 2026.005): regime regular, opcao pendente e atividade padrao (cAtivSN).
'Le e grava direto na tabela empresa (script 196). As regras de decisao ficam em modSimplesNacional.
Private mCRT As Long
Private mCarregando As Boolean

Private Sub Form_Load()
    Dim r As ADODB.Recordset
    Dim vAtiv As Variant, i As Long, sPadrao As String
    On Error GoTo erro
    mCarregando = True

    cboPendente.AddItem "Nenhuma"
    cboPendente.AddItem "Pelo Simples Nacional (CRT 1)"
    cboPendente.AddItem "Pelo Simples Nacional - excesso de sublimite (CRT 2)"
    cboPendente.ListIndex = 0

    For Each vAtiv In Array("01", "02", "03", "04", "90")
        cboAtiv.AddItem CStr(vAtiv) & " - " & SN_DescricaoCAtivSN(CStr(vAtiv))
    Next vAtiv
    cboAtiv.ListIndex = 0

    Set r = dbData.OpenRecordset("SELECT CRT, OptaRegimeRegularIBSCBS, SimplesOpcaoPendente, cAtivSNPadrao FROM empresa")
    If r Is Nothing Then
        mCRT = 0
    ElseIf r.EOF Then
        mCRT = 0
    Else
        mCRT = SN_CRT(r("CRT"))
        If Not IsNull(r("OptaRegimeRegularIBSCBS")) Then
            If CBool(r("OptaRegimeRegularIBSCBS")) Then chkRegular.Value = vbChecked
        End If
        If Not IsNull(r("SimplesOpcaoPendente")) Then
            i = Val(r("SimplesOpcaoPendente") & "")
            If i >= 0 And i <= 2 Then cboPendente.ListIndex = i
        End If
        sPadrao = SN_ResolverCAtivSN("", r("cAtivSNPadrao"))
        For i = 0 To cboAtiv.ListCount - 1
            If Left$(cboAtiv.List(i), 2) = sPadrao Then cboAtiv.ListIndex = i
        Next i
    End If

    Select Case mCRT
        Case 1: lblCRT.Caption = "CRT 1 - Simples Nacional"
        Case 2: lblCRT.Caption = "CRT 2 - Simples Nacional, excesso de sublimite"
        Case 3: lblCRT.Caption = "CRT 3 - Regime Normal"
        Case 4: lblCRT.Caption = "CRT 4 - Microempreendedor Individual (MEI)"
        Case Else: lblCRT.Caption = "CRT nao definido no cadastro da empresa"
    End Select
    mCarregando = False
    AtualizarTela
    Exit Sub
erro:
    Dim sErro As String
    sErro = Err.Number & " - " & Err.Description
    mCarregando = False
    MsgBox "Nao foi possivel carregar o enquadramento." & vbCrLf & "Confira se o script 196 (Alter_SimplesNacional_NT2026005) foi executado." & vbCrLf & vbCrLf & sErro, vbExclamation, "Aviso do Sistema"
    mCRT = 0
    AtualizarTela
End Sub

Private Sub AtualizarTela()
    Dim bReg As Boolean, nPend As Long, nTp As Long, s As String
    Dim bAnt As Boolean
    bAnt = mCarregando
    mCarregando = True

    Select Case mCRT
        Case 1, 2
            cboPendente.ListIndex = 0
            cboPendente.Enabled = False
            lblPendente.Enabled = False
            chkRegular.Enabled = True
        Case 3
            cboPendente.Enabled = True
            lblPendente.Enabled = True
            chkRegular.Enabled = (cboPendente.ListIndex > 0)
            If cboPendente.ListIndex = 0 Then chkRegular.Value = vbUnchecked
        Case Else
            cboPendente.ListIndex = 0
            cboPendente.Enabled = False
            lblPendente.Enabled = False
            chkRegular.Value = vbUnchecked
            chkRegular.Enabled = False
    End Select
    cmdSalvar.Enabled = (mCRT >= 1 And mCRT <= 3)
    cmdRBT12.Enabled = (mCRT >= 1 And mCRT <= 3)
    fraAtiv.Enabled = (mCRT >= 1 And mCRT <= 3)
    cboAtiv.Enabled = fraAtiv.Enabled

    bReg = (chkRegular.Value = vbChecked)
    nPend = cboPendente.ListIndex
    nTp = SN_TipoRegimeIBSCBS(mCRT, bReg, nPend)

    If mCRT = 4 Then
        s = "MEI: o grupo de IBS/CBS nao e informado na nota."
    ElseIf mCRT < 1 Or mCRT > 4 Then
        s = "Defina o CRT no cadastro da empresa para configurar o enquadramento."
    ElseIf SN_RegimeNormalPuro(mCRT, nTp) Then
        s = "Regime Normal: IBS/CBS informado como sempre foi; cAtivSN nao e enviado."
    Else
        If nTp = 0 Then
            s = "Nao informa tpRegimeIBSCBS. "
        Else
            s = "tpRegimeIBSCBS = " & nTp & ". "
        End If
        s = s & "cAtivSN em cada item. "
        If SN_InformaGIBSCBS(mCRT, nTp) Then
            If SN_CBSZeradaNoGIBSCBS(mCRT, nTp) Then
                s = s & "gIBSCBS com CBS zerada. "
            Else
                s = s & "gIBSCBS informado. "
            End If
        Else
            s = s & "Sem gIBSCBS. "
        End If
        If SN_InformaGTribSN(mCRT, nTp) Then
            If SN_InformaIBSNoTribSN(mCRT, nTp) Then
                s = s & "gTribSN com IBS e CBS."
            Else
                s = s & "gTribSN so com CBS."
            End If
        Else
            s = s & "Sem gTribSN."
        End If
    End If
    lblResumo.Caption = s
    mCarregando = bAnt
End Sub

Private Sub chkRegular_Click()
    If Not mCarregando Then AtualizarTela
End Sub

Private Sub cboPendente_Click()
    If Not mCarregando Then AtualizarTela
End Sub

Private Sub cmdSalvar_Click()
    Dim sSQL As String, nPend As Long, bReg As Boolean, sAtiv As String
    On Error GoTo erro
    nPend = 0
    If mCRT = 3 Then nPend = cboPendente.ListIndex
    bReg = (chkRegular.Value = vbChecked)
    If mCRT = 3 And nPend = 0 Then bReg = False
    If cboAtiv.ListIndex < 0 Then cboAtiv.ListIndex = 0
    sAtiv = Left$(cboAtiv.Text, 2)

    sSQL = "UPDATE empresa SET OptaRegimeRegularIBSCBS = " & Abs(CInt(bReg)) & _
           ", SimplesOpcaoPendente = " & nPend & _
           ", cAtivSNPadrao = '" & sAtiv & "'"
    If dbData.Execute(sSQL) Then
        MsgBox "Enquadramento salvo.", vbInformation, "Aviso"
        Unload Me
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso do Sistema"
End Sub

Private Sub cmdRBT12_Click()
    Empresa_RBT12Simples.Show vbModal
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
