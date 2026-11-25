VERSION 5.00
Begin VB.Form Empresa_RBT12Simples 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RBT12 e aliquotas do Simples Nacional (IBS/CBS)"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "RBT12 = receita bruta dos 12 meses anteriores ao mes da apuracao (o valor do PGDAS-D e a fonte oficial). Usado na emissao a partir de 01/01/2027."
      ForeColor       =   &H00606060&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7320
   End
   Begin VB.Frame fraRBT 
      Caption         =   "RBT12 da competencia"
      Height          =   2475
      Left            =   240
      TabIndex        =   1
      Top             =   780
      Width           =   7320
      Begin VB.Label lblComp 
         BackStyle       =   0  'Transparent
         Caption         =   "Competencia (MM/AAAA):"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   375
         Width           =   2000
      End
      Begin VB.TextBox txtComp 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lblRBT 
         BackStyle       =   0  'Transparent
         Caption         =   "RBT12 (R$):"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   825
         Width           =   2000
      End
      Begin VB.TextBox txtRBT12 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   780
         Width           =   2400
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular pelo sistema"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   750
         Width           =   2200
      End
      Begin VB.Label lblOrigem 
         BackStyle       =   0  'Transparent
         Caption         =   ""
         ForeColor       =   &H00804000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1260
         Width           =   6840
      End
      Begin VB.Label lblPrev 
         BackStyle       =   0  'Transparent
         Caption         =   ""
         ForeColor       =   &H00804000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   6840
      End
   End
   Begin VB.Frame fraAnexo 
      Caption         =   "Anexo padrao da empresa"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   7320
      Begin VB.ComboBox cboAnexo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   6840
      End
   End
   Begin VB.Frame fraOver 
      Caption         =   "Percentuais do PGDAS-D (opcional - se preenchidos, valem no lugar do calculo)"
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   7320
      Begin VB.Label lblH1 
         BackStyle       =   0  'Transparent
         Caption         =   "IBS (pIBSSN %)"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblH2 
         BackStyle       =   0  'Transparent
         Caption         =   "CBS (pCBSSN %)"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo I - Comercio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   665
         Width           =   3200
      End
      Begin VB.TextBox txtPIBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   3600
         TabIndex        =   15
         Top             =   620
         Width           =   1500
      End
      Begin VB.TextBox txtPCBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   5400
         TabIndex        =   16
         Top             =   620
         Width           =   1500
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo II - Industria"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1035
         Width           =   3200
      End
      Begin VB.TextBox txtPIBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   18
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox txtPCBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   5400
         TabIndex        =   19
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo III - Locacao e servicos"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1405
         Width           =   3200
      End
      Begin VB.TextBox txtPIBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3600
         TabIndex        =   21
         Top             =   1360
         Width           =   1500
      End
      Begin VB.TextBox txtPCBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5400
         TabIndex        =   22
         Top             =   1360
         Width           =   1500
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo IV - Servicos"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   1775
         Width           =   3200
      End
      Begin VB.TextBox txtPIBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3600
         TabIndex        =   24
         Top             =   1730
         Width           =   1500
      End
      Begin VB.TextBox txtPCBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   5400
         TabIndex        =   25
         Top             =   1730
         Width           =   1500
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo V - Servicos (fator r < 28%)"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   2145
         Width           =   3200
      End
      Begin VB.TextBox txtPIBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   3600
         TabIndex        =   27
         Top             =   2100
         Width           =   1500
      End
      Begin VB.TextBox txtPCBS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   5400
         TabIndex        =   28
         Top             =   2100
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Default         =   -1  'True
      Height          =   435
      Left            =   4200
      TabIndex        =   29
      Top             =   6900
      Width           =   1575
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   435
      Left            =   5880
      TabIndex        =   30
      Top             =   6900
      Width           =   1680
   End
End
Attribute VB_Name = "Empresa_RBT12Simples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RBT12 por competencia (AAAAMM) e anexo padrao do Simples Nacional (NT 2026.005).
'Grava em SimplesNacionalApuracao / SimplesNacionalApuracaoAnexo / Empresa.AnexoSNPadrao (script 198).
'O calculo das aliquotas fica em modSimplesNacional (script 197).
Private mCRT As Long
Private mCarregando As Boolean
Private mOrigem As String
Private mCompCarregada As String

Private Sub Form_Load()
    Dim r As ADODB.Recordset
    Dim i As Long, sAnexo As String
    On Error GoTo erro
    mCarregando = True
    mOrigem = "MANUAL"
    cboAnexo.AddItem "I - Comercio"
    cboAnexo.AddItem "II - Industria"
    cboAnexo.AddItem "III - Locacao e servicos"
    cboAnexo.AddItem "IV - Servicos"
    cboAnexo.AddItem "V - Servicos (fator r < 28%)"
    cboAnexo.ListIndex = 0
    txtComp.Text = Format$(Date, "mm/yyyy")

    Set r = dbData.OpenRecordset("SELECT CRT, AnexoSNPadrao FROM empresa")
    If r Is Nothing Then
        mCRT = 0
    ElseIf r.EOF Then
        mCRT = 0
    Else
        mCRT = SN_CRT(r("CRT"))
        sAnexo = Trim$(r("AnexoSNPadrao") & "")
        For i = 0 To 4
            If Split(cboAnexo.List(i), " ")(0) = sAnexo Then cboAnexo.ListIndex = i
        Next i
    End If
    mCarregando = False
    CarregarCompetencia
    Exit Sub
erro:
    Dim sErro As String
    sErro = Err.Number & " - " & Err.Description
    mCarregando = False
    MsgBox "Nao foi possivel carregar os dados." & vbCrLf & "Confira se o script 198 (Alter_SimplesNacionalApuracao) foi executado." & vbCrLf & vbCrLf & sErro, vbExclamation, "Aviso do Sistema"
End Sub

'Texto digitado -> numero (aceita 1.234.567,89)
Private Function Num(ByVal s As String) As Double
    s = Replace(Trim$(s), ".", "")
    s = Replace(s, ",", ".")
    Num = Val(s)
End Function

Private Function SQLN(ByVal d As Double) As String
    SQLN = Replace(CStr(d), ",", ".")
End Function

'MM/AAAA -> AAAAMM ("" se invalida)
Private Function CompAAAAMM(ByVal s As String) As String
    Dim m As Long, a As Long
    CompAAAAMM = ""
    s = Trim$(s)
    If Len(s) <> 7 Then Exit Function
    If Mid$(s, 3, 1) <> "/" Then Exit Function
    m = Val(Left$(s, 2)): a = Val(Right$(s, 4))
    If m < 1 Or m > 12 Or a < 2000 Or a > 2099 Then Exit Function
    CompAAAAMM = Format$(a, "0000") & Format$(m, "00")
End Function

Private Function AnexoSelecionado() As String
    If cboAnexo.ListIndex < 0 Then
        AnexoSelecionado = "I"
    Else
        AnexoSelecionado = Split(cboAnexo.List(cboAnexo.ListIndex), " ")(0)
    End If
End Function

'Carrega a competencia digitada; se nao existir, mostra a mais recente anterior (nao trava a emissao)
Private Sub CarregarCompetencia()
    Dim r As ADODB.Recordset, sComp As String, i As Long
    On Error GoTo erro
    sComp = CompAAAAMM(txtComp.Text)
    If sComp = "" Then Exit Sub
    mCarregando = True
    txtRBT12.Text = ""
    For i = 0 To 4
        txtPIBS(i).Text = ""
        txtPCBS(i).Text = ""
    Next i
    lblOrigem.Caption = "Competencia sem RBT12 cadastrado."
    mOrigem = "MANUAL"

    Set r = dbData.OpenRecordset("SELECT TOP 1 Competencia, RBT12, Origem FROM SimplesNacionalApuracao WHERE Competencia <= '" & sComp & "' ORDER BY Competencia DESC")
    If Not r Is Nothing Then
        If Not r.EOF Then
            txtRBT12.Text = Format$(CDbl(r("RBT12")), "#,##0.00")
            mOrigem = Trim$(r("Origem") & "")
            If Trim$(r("Competencia") & "") = sComp Then
                lblOrigem.Caption = "RBT12 gravado nesta competencia (origem: " & mOrigem & ")."
            Else
                lblOrigem.Caption = "Sem valor para esta competencia: repetindo o de " & Right$(Trim$(r("Competencia") & ""), 2) & "/" & Left$(Trim$(r("Competencia") & ""), 4) & ". Confira com o contador."
            End If
        End If
    End If

    Set r = dbData.OpenRecordset("SELECT Anexo, pIBSSN, pCBSSN FROM SimplesNacionalApuracaoAnexo WHERE Competencia = '" & sComp & "'")
    If Not r Is Nothing Then
        Do While Not r.EOF
            Select Case Trim$(r("Anexo") & "")
                Case "I": i = 0
                Case "II": i = 1
                Case "III": i = 2
                Case "IV": i = 3
                Case "V": i = 4
                Case Else: i = -1
            End Select
            If i >= 0 Then
                If Not IsNull(r("pIBSSN")) Then txtPIBS(i).Text = Format$(CDbl(r("pIBSSN")), "0.0000")
                If Not IsNull(r("pCBSSN")) Then txtPCBS(i).Text = Format$(CDbl(r("pCBSSN")), "0.0000")
            End If
            r.MoveNext
        Loop
    End If
    mCompCarregada = sComp
    mCarregando = False
    AtualizarPrevia
    Exit Sub
erro:
    Dim sErro As String
    sErro = Err.Number & " - " & Err.Description
    mCarregando = False
    MsgBox "Erro ao ler a competencia: " & sErro, vbExclamation, "Aviso do Sistema"
End Sub

'Mostra as aliquotas que seriam usadas com o RBT12 e o anexo padrao
Private Sub AtualizarPrevia()
    Dim sComp As String, pC As Double, pI As Double, dRBT As Double
    On Error GoTo erro
    lblPrev.Caption = ""
    sComp = CompAAAAMM(txtComp.Text)
    If sComp = "" Then Exit Sub
    If mCRT < 1 Or mCRT > 3 Then
        lblPrev.Caption = "Defina o CRT no cadastro da empresa."
        Exit Sub
    End If
    dRBT = Num(txtRBT12.Text)
    If SN_PercentuaisDoAnexo(AnexoSelecionado(), CLng(Left$(sComp, 4)), mCRT, dRBT, pC, pI) Then
        lblPrev.Caption = "Anexo " & AnexoSelecionado() & " em " & Left$(sComp, 4) & ": pCBSSN = " & Format$(pC, "0.0000") & "%   pIBSSN = " & Format$(pI, "0.0000") & "%"
    Else
        If CLng(Left$(sComp, 4)) < 2027 Then
            lblPrev.Caption = "Antes de 2027 nao ha IBS/CBS no Simples Nacional: informe uma competencia a partir de 01/2027 para ver as aliquotas."
        Else
            lblPrev.Caption = "Nao ha aliquota cadastrada para este anexo/ano (script 197)."
        End If
    End If
    Exit Sub
erro:
    lblPrev.Caption = "Previa indisponivel: " & Err.Description
End Sub

Private Sub txtComp_LostFocus()
    If CompAAAAMM(txtComp.Text) = "" Then
        MsgBox "Competencia invalida. Use MM/AAAA, por exemplo 03/2027.", vbExclamation, "Aviso"
        txtComp.SetFocus
    ElseIf CompAAAAMM(txtComp.Text) <> mCompCarregada Then
        CarregarCompetencia
    End If
End Sub

Private Sub txtRBT12_Change()
    If mCarregando Then Exit Sub
    mOrigem = "MANUAL"
    AtualizarPrevia
End Sub

Private Sub txtRBT12_LostFocus()
    If Trim$(txtRBT12.Text) <> "" Then
        mCarregando = True
        txtRBT12.Text = Format$(Num(txtRBT12.Text), "#,##0.00")
        mCarregando = False
        AtualizarPrevia
    End If
End Sub

Private Sub cboAnexo_Click()
    If Not mCarregando Then AtualizarPrevia
End Sub

'Sugestao: soma NF-e (saida) e NFC-e autorizadas e nao canceladas dos 12 meses antes do mes da competencia
Private Sub cmdCalcular_Click()
    Dim r As ADODB.Recordset, sComp As String, dIni As Date, dFim As Date
    Dim dTot As Double, sIni As String, sFim As String
    On Error GoTo erro
    sComp = CompAAAAMM(txtComp.Text)
    If sComp = "" Then
        MsgBox "Informe a competencia (MM/AAAA).", vbExclamation, "Aviso"
        Exit Sub
    End If
    dFim = DateSerial(CLng(Left$(sComp, 4)), CLng(Right$(sComp, 2)), 1)
    dIni = DateAdd("m", -12, dFim)
    sIni = Format$(dIni, "yyyymmdd")
    sFim = Format$(dFim, "yyyymmdd")

    Set r = dbData.OpenRecordset("SELECT ISNULL(SUM(ValorNota), 0) AS T FROM NotaFiscal WHERE TipoDocumento = 1 AND ISNULL(Cancelada, 0) = 0 " & _
        "AND LEN(ISNULL(NumeroProtocolo, '')) > 1 AND ISNULL(FinalidadeEmissaoNFe, '') LIKE '1%' " & _
        "AND DataEmissao >= '" & sIni & "' AND DataEmissao < '" & sFim & "'")
    If Not r Is Nothing Then
        If Not r.EOF Then dTot = dTot + CDbl(r("T"))
    End If
    Set r = dbData.OpenRecordset("SELECT ISNULL(SUM(Valor_NF_Prod), 0) AS T FROM TbNFCe WHERE TipoNF = 'S' AND ISNULL(Cancelada, 0) = 0 " & _
        "AND ISNULL(NFCeCancelada, 0) = 0 AND LEN(ISNULL(NFCeProtocolo, '')) > 1 " & _
        "AND DataEmissao >= '" & sIni & "' AND DataEmissao < '" & sFim & "'")
    If Not r Is Nothing Then
        If Not r.EOF Then dTot = dTot + CDbl(r("T"))
    End If

    mCarregando = True
    txtRBT12.Text = Format$(dTot, "#,##0.00")
    mCarregando = False
    mOrigem = "SISTEMA"
    lblOrigem.Caption = "Sugestao pelo sistema: notas de " & Format$(dIni, "mm/yyyy") & " a " & Format$(DateAdd("m", -1, dFim), "mm/yyyy") & ". Nao inclui receitas de fora do sistema. Confira com o PGDAS-D."
    AtualizarPrevia
    Exit Sub
erro:
    MsgBox "Erro ao calcular: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso do Sistema"
End Sub

Private Sub cmdSalvar_Click()
    Dim sComp As String, sSQL As String, i As Long
    Dim sI As String, sC As String, vA As Variant
    On Error GoTo erro
    sComp = CompAAAAMM(txtComp.Text)
    If sComp = "" Then
        MsgBox "Competencia invalida. Use MM/AAAA.", vbExclamation, "Aviso"
        Exit Sub
    End If
    If Trim$(txtRBT12.Text) = "" Then
        MsgBox "Informe o RBT12 (ou use Calcular pelo sistema).", vbExclamation, "Aviso"
        txtRBT12.SetFocus
        Exit Sub
    End If
    If Num(txtRBT12.Text) < 0 Then
        MsgBox "RBT12 nao pode ser negativo.", vbExclamation, "Aviso"
        Exit Sub
    End If

    sSQL = "SET XACT_ABORT ON; BEGIN TRANSACTION; " & _
           "IF EXISTS (SELECT 1 FROM SimplesNacionalApuracao WHERE Competencia = '" & sComp & "') " & _
           "UPDATE SimplesNacionalApuracao SET RBT12 = " & SQLN(Num(txtRBT12.Text)) & ", Origem = '" & mOrigem & "', DataGravacao = GETDATE() WHERE Competencia = '" & sComp & "' " & _
           "ELSE INSERT INTO SimplesNacionalApuracao (Competencia, RBT12, Origem) VALUES ('" & sComp & "', " & SQLN(Num(txtRBT12.Text)) & ", '" & mOrigem & "'); " & _
           "DELETE FROM SimplesNacionalApuracaoAnexo WHERE Competencia = '" & sComp & "'; "
    vA = Array("I", "II", "III", "IV", "V")
    For i = 0 To 4
        If Trim$(txtPIBS(i).Text) <> "" Or Trim$(txtPCBS(i).Text) <> "" Then
            sI = "NULL": sC = "NULL"
            If Trim$(txtPIBS(i).Text) <> "" Then sI = SQLN(Num(txtPIBS(i).Text))
            If Trim$(txtPCBS(i).Text) <> "" Then sC = SQLN(Num(txtPCBS(i).Text))
            sSQL = sSQL & "INSERT INTO SimplesNacionalApuracaoAnexo (Competencia, Anexo, pIBSSN, pCBSSN) VALUES ('" & sComp & "', '" & CStr(vA(i)) & "', " & sI & ", " & sC & "); "
        End If
    Next i
    sSQL = sSQL & "UPDATE empresa SET AnexoSNPadrao = '" & AnexoSelecionado() & "'; COMMIT TRANSACTION;"
    If dbData.Execute(sSQL) Then
        MsgBox "RBT12 e aliquotas salvos.", vbInformation, "Aviso"
        Unload Me
    End If
    Exit Sub
erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Aviso do Sistema"
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
