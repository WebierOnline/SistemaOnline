VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSPED_Conferencia
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPED Fiscal - Conferência dos cadastros"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   12720
      TabIndex        =   6
      Top             =   6720
      Width           =   1500
   End
   Begin VB.CommandButton cmdGerar
      Caption         =   "&Gerar o arquivo"
      Height          =   420
      Left            =   10920
      TabIndex        =   5
      Top             =   6720
      Width           =   1740
   End
   Begin VB.CommandButton cmdSalvar
      Caption         =   "&Salvar lista"
      Height          =   420
      Left            =   4260
      TabIndex        =   4
      Top             =   6720
      Width           =   1500
   End
   Begin VB.CommandButton cmdMarcar
      Caption         =   "&Marcar / desmarcar"
      Height          =   420
      Left            =   2280
      TabIndex        =   3
      Top             =   6720
      Width           =   1920
   End
   Begin VB.CommandButton cmdCorrigir
      Caption         =   "C&orrigir os marcados"
      Height          =   420
      Left            =   60
      TabIndex        =   2
      Top             =   6720
      Width           =   2160
   End
   Begin MSFlexGridLib.MSFlexGrid grdLista
      Height          =   5955
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10504
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label lblResumo
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   14175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSPED_Conferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Conferência dos cadastros antes de gerar o SPED (chamada pelo frmSPED_Fiscal): lista o que o SpedGerador achou
' (modResumoFiscal.SPED_Conferir). "Corrige" = há fonte confiável (XML da nota de entrada ou tabela de cidades) e vem
' marcado; "Corrigir à mão" = sem fonte (ou NF-e já emitida, que não é corrigida); "Conferir" = NCM válido mas
' diferente da nota de entrada. As correções guardam o valor anterior em SPED_CorrecaoCadastro.
Private Const COL_MARCA As Integer = 0
Private Const COL_SITUACAO As Integer = 1
Private Const COL_TIPO As Integer = 2
Private Const COL_CODIGO As Integer = 3
Private Const COL_NOME As Integer = 4
Private Const COL_CAMPO As Integer = 5
Private Const COL_ATUAL As Integer = 6
Private Const COL_NOVO As Integer = 7
Private Const COL_PROBLEMA As Integer = 8
Private Const COL_FONTE As Integer = 9
Private Const COL_ONDE As Integer = 10
Private Const COL_CHAVE As Integer = 11
Private Const COL_SIT As Integer = 12

Dim sParametros As String
Dim bGerar As Boolean

' Mostra a conferência (modal). True = gerar o arquivo; False = cancelado.
Public Function Mostrar(ByVal sParam As String) As Boolean
    sParametros = sParam
    bGerar = False
    MontarGrade
    Carregar
    Me.Show vbModal
    Mostrar = bGerar
    Unload Me
End Function

Private Sub MontarGrade()
    With grdLista
        .Cols = 13
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColWidth(COL_MARCA) = 450: .TextMatrix(0, COL_MARCA) = "Corr."
        .ColWidth(COL_SITUACAO) = 1300: .TextMatrix(0, COL_SITUACAO) = "Situação"
        .ColWidth(COL_TIPO) = 1100: .TextMatrix(0, COL_TIPO) = "Tipo"
        .ColWidth(COL_CODIGO) = 750: .TextMatrix(0, COL_CODIGO) = "Código"
        .ColWidth(COL_NOME) = 2600: .TextMatrix(0, COL_NOME) = "Nome"
        .ColWidth(COL_CAMPO) = 850: .TextMatrix(0, COL_CAMPO) = "Campo"
        .ColWidth(COL_ATUAL) = 1000: .TextMatrix(0, COL_ATUAL) = "Atual"
        .ColWidth(COL_NOVO) = 1000: .TextMatrix(0, COL_NOVO) = "Correção"
        .ColWidth(COL_PROBLEMA) = 2600: .TextMatrix(0, COL_PROBLEMA) = "Problema"
        .ColWidth(COL_FONTE) = 1900: .TextMatrix(0, COL_FONTE) = "Fonte da correção"
        .ColWidth(COL_ONDE) = 2200: .TextMatrix(0, COL_ONDE) = "Onde corrigir"
        .ColWidth(COL_CHAVE) = 0
        .ColWidth(COL_SIT) = 0
    End With
End Sub

Private Sub Carregar()
    Dim nArq As Integer
    Dim sLinha As String
    Dim v As Variant
    Dim r As Long, c As Integer
    Dim nCorr As Long, nMan As Long, nDiv As Long

    grdLista.Rows = 1
    If Dir$(SPED_ArquivoConferencia) = "" Then GoTo resumo

    nArq = FreeFile
    Open SPED_ArquivoConferencia For Input As #nArq
    Do While Not EOF(nArq)
        Line Input #nArq, sLinha
        v = Split(sLinha, vbTab)
        If UBound(v) >= 10 Then
            grdLista.AddItem ""
            r = grdLista.Rows - 1
            grdLista.TextMatrix(r, COL_SIT) = v(1)
            Select Case v(1)
                Case "CORRIGIR"
                    grdLista.TextMatrix(r, COL_SITUACAO) = "Corrige"
                    grdLista.TextMatrix(r, COL_MARCA) = "X"
                    nCorr = nCorr + 1
                Case "DIVERGENTE"
                    grdLista.TextMatrix(r, COL_SITUACAO) = "Conferir"
                    nDiv = nDiv + 1
                Case Else
                    grdLista.TextMatrix(r, COL_SITUACAO) = "Corrigir à mão"
                    nMan = nMan + 1
            End Select
            grdLista.TextMatrix(r, COL_CHAVE) = v(0)
            grdLista.TextMatrix(r, COL_TIPO) = v(2)
            grdLista.TextMatrix(r, COL_CODIGO) = v(3)
            grdLista.TextMatrix(r, COL_NOME) = v(4)
            grdLista.TextMatrix(r, COL_CAMPO) = v(5)
            grdLista.TextMatrix(r, COL_ATUAL) = v(6)
            grdLista.TextMatrix(r, COL_NOVO) = v(7)
            grdLista.TextMatrix(r, COL_PROBLEMA) = v(8)
            grdLista.TextMatrix(r, COL_FONTE) = v(9)
            grdLista.TextMatrix(r, COL_ONDE) = v(10)
            'cor da linha: verde = corrige sozinho, amarelo = conferir, vermelho = corrigir à mão
            grdLista.Row = r
            For c = 0 To COL_ONDE
                grdLista.Col = c
                Select Case v(1)
                    Case "CORRIGIR": grdLista.CellBackColor = RGB(220, 245, 220)
                    Case "DIVERGENTE": grdLista.CellBackColor = RGB(255, 245, 200)
                    Case Else: grdLista.CellBackColor = RGB(250, 220, 220)
                End Select
            Next
        End If
    Loop
    Close #nArq

resumo:
    If nCorr + nMan + nDiv = 0 Then
        lblResumo.Caption = "Nenhuma pendência nos cadastros que vão no arquivo."
    Else
        lblResumo.Caption = nCorr & " cadastro(s) podem ser corrigidos pelo sistema (verde, marcados com X - clique na coluna Corr. para desmarcar), " & _
                            nMan & " precisam ser corrigidos à mão (vermelho) e " & nDiv & " estão diferentes da nota de entrada (amarelo, confira qual está certo)." & vbCrLf & _
                            "Notas fiscais (NF-e/NFC-e) já emitidas não são corrigidas."
    End If
    cmdCorrigir.Enabled = (nCorr > 0)
    cmdMarcar.Enabled = (nCorr > 0)
    cmdSalvar.Enabled = (nCorr + nMan + nDiv > 0)
End Sub

Private Sub grdLista_Click()
    Dim r As Long
    r = grdLista.MouseRow
    If r < 1 Or grdLista.MouseCol <> COL_MARCA Then Exit Sub
    If grdLista.TextMatrix(r, COL_SIT) <> "CORRIGIR" Then Exit Sub
    grdLista.TextMatrix(r, COL_MARCA) = IIf(grdLista.TextMatrix(r, COL_MARCA) = "X", "", "X")
End Sub

Private Sub cmdMarcar_Click()
    Dim r As Long, sMarca As String
    sMarca = "X"
    For r = 1 To grdLista.Rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = "CORRIGIR" And grdLista.TextMatrix(r, COL_MARCA) = "X" Then sMarca = "": Exit For
    Next
    For r = 1 To grdLista.Rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = "CORRIGIR" Then grdLista.TextMatrix(r, COL_MARCA) = sMarca
    Next
End Sub

Private Sub cmdCorrigir_Click()
    Dim r As Long, n As Long, nFeitos As Long
    Dim sChaves As String, sRes As String
    Dim nC As Long, nM As Long, nD As Long

    For r = 1 To grdLista.Rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = "CORRIGIR" And grdLista.TextMatrix(r, COL_MARCA) = "X" Then
            sChaves = sChaves & IIf(sChaves <> "", ",", "") & grdLista.TextMatrix(r, COL_CHAVE)
            n = n + 1
        End If
    Next
    If n = 0 Then
        MsgBox "Marque (X) os cadastros a corrigir.", vbInformation
        Exit Sub
    End If
    If MsgBox("Corrigir " & n & " cadastro(s) com os valores da coluna Correção?" & vbCrLf & vbCrLf & _
              "O valor anterior fica guardado no histórico (tabela SPED_CorrecaoCadastro).", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    DoEvents
    sRes = SPED_Corrigir(sParametros, sChaves, nFeitos)
    If sRes = "OK" Then sRes = SPED_Conferir(sParametros, nC, nM, nD)
    Screen.MousePointer = vbDefault
    If sRes <> "OK" Then
        MsgBox "Não foi possível corrigir: " & Mid$(sRes, IIf(Left$(sRes, 6) = "ERRO: ", 7, 1)), vbExclamation
        Exit Sub
    End If
    Carregar
    MsgBox nFeitos & " cadastro(s) corrigido(s).", vbInformation
End Sub

' Lista em texto (para corrigir os cadastros com calma ou mandar para alguém), aberta no Bloco de Notas.
Private Sub cmdSalvar_Click()
    Dim sPasta As String, sArq As String
    Dim nArq As Integer, r As Long
    Dim oFso As Object

    On Error GoTo ErrL
    sPasta = SPED_Pasta
    Set oFso = CreateObject("Scripting.FileSystemObject")
    If Not oFso.FolderExists(oFso.GetParentFolderName(sPasta)) Then oFso.CreateFolder oFso.GetParentFolderName(sPasta)
    If Not oFso.FolderExists(sPasta) Then oFso.CreateFolder sPasta
    sArq = sPasta & "\Conferencia_cadastros_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    nArq = FreeFile
    Open sArq For Output As #nArq
    Print #nArq, "SPED Fiscal - conferência dos cadastros - " & Format(Now, "dd/mm/yyyy hh:nn")
    Print #nArq, String(100, "-")
    For r = 1 To grdLista.Rows - 1
        Print #nArq, grdLista.TextMatrix(r, COL_SITUACAO) & " | " & grdLista.TextMatrix(r, COL_TIPO) & " " & grdLista.TextMatrix(r, COL_CODIGO) & _
                     " - " & grdLista.TextMatrix(r, COL_NOME)
        Print #nArq, "   " & grdLista.TextMatrix(r, COL_CAMPO) & ": " & grdLista.TextMatrix(r, COL_PROBLEMA) & _
                     " (atual: " & grdLista.TextMatrix(r, COL_ATUAL) & IIf(grdLista.TextMatrix(r, COL_NOVO) <> "", "; correção: " & grdLista.TextMatrix(r, COL_NOVO), "") & ")"
        If grdLista.TextMatrix(r, COL_FONTE) <> "" Then Print #nArq, "   Fonte: " & grdLista.TextMatrix(r, COL_FONTE)
        Print #nArq, "   Onde corrigir: " & grdLista.TextMatrix(r, COL_ONDE)
    Next
    Close #nArq
    Shell "notepad.exe " & Chr(34) & sArq & Chr(34), vbNormalFocus
    Exit Sub
ErrL:
    MsgBox "Não foi possível salvar a lista: " & Err.Description, vbExclamation
End Sub

Private Sub cmdGerar_Click()
    Dim r As Long, n As Long
    For r = 1 To grdLista.Rows - 1
        If grdLista.TextMatrix(r, COL_SIT) = "MANUAL" Or grdLista.TextMatrix(r, COL_SIT) = "CORRIGIR" Then n = n + 1
    Next
    If n > 0 Then
        If MsgBox("Ainda há " & n & " pendência(s) que o validador do SPED deve apontar." & vbCrLf & "Gerar o arquivo assim mesmo?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    bGerar = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    bGerar = False
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'fechar pelo X = cancelar (a função Mostrar descarrega o form)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        bGerar = False
        Me.Hide
    End If
End Sub
