VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCBS_Aliquotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CBS - Alíquotas por Vigência"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGrade 
      Caption         =   "Alíquotas cadastradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid lstCBS 
         Height          =   2820
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   4974
         _Version        =   393216
      End
   End
   Begin VB.Frame fraEdicao 
      Caption         =   "Edição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   60
      TabIndex        =   2
      Top             =   3300
      Width           =   9135
      Begin VB.TextBox txtObs 
         Height          =   288
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   9
         Top             =   780
         Width           =   7740
      End
      Begin VB.TextBox txtFim 
         Height          =   288
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   7
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txtIni 
         Height          =   288
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txtAliq 
         Height          =   288
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   3
         Top             =   300
         Width           =   900
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   360
         Left            =   3000
         TabIndex        =   12
         Top             =   1260
         Width           =   1440
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   360
         Left            =   1500
         TabIndex        =   11
         Top             =   1260
         Width           =   1440
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   1260
         Width           =   1320
      End
      Begin VB.Label lblObs 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Width           =   885
      End
      Begin VB.Label lblFim 
         AutoSize        =   -1  'True
         Caption         =   "Fim (vazio = sem fim):"
         Height          =   195
         Left            =   4440
         TabIndex        =   6
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblIni 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   300
         Width           =   465
      End
      Begin VB.Label lblAliq 
         AutoSize        =   -1  'True
         Caption         =   "Alíq. CBS %:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   360
      Left            =   7755
      TabIndex        =   14
      Top             =   6300
      Width           =   1440
   End
   Begin VB.Label lblVigente 
      Height          =   420
      Left            =   120
      TabIndex        =   15
      Top             =   5220
      Width           =   9000
   End
   Begin VB.Label lblAviso 
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   120
      TabIndex        =   16
      Top             =   5700
      Width           =   7500
   End
End
Attribute VB_Name = "frmCBS_Aliquotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iId As Long

' Cadastro da alíquota da CBS por vigência (tabela CBS_Aliquota, script 168).
' A alíquota vigente hoje é copiada para Empresa.CBSpAliq por AplicarAliquotasReforma (modNFe),
' que roda antes de calcular IBS/CBS de cada nota - a virada de vigência é automática.
Private Sub Form_Load()
    On Error GoTo ErrL
    With lstCBS
        .Cols = 5
        .FixedCols = 0
        .ColWidth(0) = 1300: .Row = 0: .Col = 0: .Text = "Alíq. %"
        .ColWidth(1) = 1500: .Row = 0: .Col = 1: .Text = "Vig. Início"
        .ColWidth(2) = 1500: .Row = 0: .Col = 2: .Text = "Vig. Fim"
        .ColWidth(3) = 4500: .Row = 0: .Col = 3: .Text = "Observação"
        .ColWidth(4) = 0     ' Id oculto
        .AllowUserResizing = 1: .SelectionMode = 1
    End With
    Carregar
    Exit Sub
ErrL:
    MsgBox "Não foi possível abrir as alíquotas da CBS." & vbCrLf & vbCrLf & _
           "Verifique se o script AlterTable_AliquotasReforma_Vigencia já foi aplicado no banco." & vbCrLf & _
           Err.Description, vbExclamation
    Unload Me
End Sub

Private Sub Carregar()
    Dim rRs As ADODB.Recordset
    Dim r As Long
    Dim sVigente As String
    lstCBS.rows = 1
    RsOpen rRs, "SELECT Id, CBSpAliq, dIniVig, dFimVig, Obs FROM CBS_Aliquota ORDER BY dIniVig"
    Do While Not rRs.EOF
        lstCBS.AddItem ""
        r = lstCBS.rows - 1
        lstCBS.TextMatrix(r, 0) = Format(rRs("CBSpAliq"), "0.00##")
        lstCBS.TextMatrix(r, 1) = Format(rRs("dIniVig"), "dd/mm/yyyy")
        lstCBS.TextMatrix(r, 2) = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "dd/mm/yyyy"))
        lstCBS.TextMatrix(r, 3) = rRs("Obs") & ""
        lstCBS.TextMatrix(r, 4) = rRs("Id")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    Novo
    MostrarVigente
End Sub

Private Sub MostrarVigente()
    Dim sAliq As String, sEmp As String, sHoje As String
    sHoje = Format(Date, "yyyymmdd")
    sAliq = SQLExecutaRetorno("SELECT TOP 1 CAST(CBSpAliq AS VARCHAR(20)) AS A FROM CBS_Aliquota WHERE '" & sHoje & "' BETWEEN dIniVig AND ISNULL(dFimVig,'99991231') ORDER BY dIniVig DESC", "A", "")
    sEmp = SQLExecutaRetorno("SELECT TOP 1 CAST(CBSpAliq AS VARCHAR(20)) AS A FROM Empresa", "A", "")
    If sAliq = "" Then
        lblVigente.Caption = "Nenhuma alíquota cadastrada cobre a data de hoje (" & Format(Date, "dd/mm/yyyy") & "). O sistema continua usando o valor da empresa: " & sEmp & "%."
        lblAviso.Caption = "ATENÇÃO: sem alíquota vigente, a CBS das notas não muda sozinha."
    Else
        lblVigente.Caption = "Vigente hoje (" & Format(Date, "dd/mm/yyyy") & "): " & Format(Val(sAliq), "0.00##") & "%.   Gravado na empresa: " & sEmp & "%."
        lblAviso.Caption = ""
    End If
End Sub

Private Sub lstCBS_Click()
    If lstCBS.Row = 0 Then Exit Sub
    iId = Val(lstCBS.TextMatrix(lstCBS.Row, 4))
    txtAliq.Text = lstCBS.TextMatrix(lstCBS.Row, 0)
    txtIni.Text = lstCBS.TextMatrix(lstCBS.Row, 1)
    txtFim.Text = lstCBS.TextMatrix(lstCBS.Row, 2)
    txtObs.Text = lstCBS.TextMatrix(lstCBS.Row, 3)
End Sub

Private Sub Novo()
    iId = 0
    txtAliq.Text = "": txtIni.Text = "": txtFim.Text = "": txtObs.Text = ""
End Sub

Private Sub cmdNovo_Click()
    Novo
    txtAliq.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    Dim dAliq As Double, sAliq As String
    Dim sIni As String, sFim As String, sObs As String
    Dim sFimSQL As String, sql As String, sConflito As String

    dAliq = Val(Replace(Trim(txtAliq.Text), ",", "."))
    sIni = Trim(txtIni.Text)
    sFim = Trim(txtFim.Text)
    sObs = Left$(Replace(Trim(txtObs.Text), "'", "''"), 200)

    If dAliq <= 0 Or dAliq > 100 Then
        MsgBox "Informe a alíquota da CBS em % (ex.: 0,90 ou 8,80).", vbExclamation: txtAliq.SetFocus: Exit Sub
    End If
    If Not IsDate(sIni) Then
        MsgBox "Informe o início da vigência no formato dd/mm/aaaa.", vbExclamation: txtIni.SetFocus: Exit Sub
    End If
    If sFim <> "" Then
        If Not IsDate(sFim) Then
            MsgBox "Fim da vigência inválido. Use dd/mm/aaaa ou deixe vazio (sem fim).", vbExclamation: txtFim.SetFocus: Exit Sub
        End If
        If CDate(sFim) < CDate(sIni) Then
            MsgBox "O fim da vigência não pode ser anterior ao início.", vbExclamation: txtFim.SetFocus: Exit Sub
        End If
    End If
    sIni = Format(CDate(sIni), "dd/mm/yyyy")
    If sFim <> "" Then sFim = Format(CDate(sFim), "dd/mm/yyyy")
    sAliq = Replace(CStr(dAliq), ",", ".")
    sFimSQL = IIf(sFim = "", "'99991231'", "CONVERT(date,'" & sFim & "',103)")

    ' vigências não podem se sobrepor
    sConflito = SQLExecutaRetorno("SELECT CAST(COUNT(*) AS VARCHAR) AS T FROM CBS_Aliquota WHERE Id <> " & iId & _
                " AND dIniVig <= " & sFimSQL & " AND ISNULL(dFimVig,'99991231') >= CONVERT(date,'" & sIni & "',103)", "T", "0")
    If Val(sConflito) > 0 Then
        MsgBox "Este período se sobrepõe a outra alíquota já cadastrada." & vbCrLf & _
               "Ajuste o fim da vigência anterior (ex.: 31/12/2026) antes de salvar.", vbExclamation
        Exit Sub
    End If

    sFimSQL = IIf(sFim = "", "NULL", "CONVERT(date,'" & sFim & "',103)")
    If iId = 0 Then
        sql = "INSERT INTO CBS_Aliquota (CBSpAliq, dIniVig, dFimVig, Obs) VALUES (" & sAliq & ",CONVERT(date,'" & sIni & "',103)," & sFimSQL & ",'" & sObs & "')"
    Else
        sql = "UPDATE CBS_Aliquota SET CBSpAliq=" & sAliq & ",dIniVig=CONVERT(date,'" & sIni & "',103),dFimVig=" & sFimSQL & ",Obs='" & sObs & "' WHERE Id=" & iId
    End If
    On Error GoTo ErrS
    vgDb.Execute sql
    AplicarAliquotasReforma
    MsgBox "Salvo.", vbInformation
    Carregar
    Exit Sub
ErrS:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluir_Click()
    If iId = 0 Then MsgBox "Selecione uma alíquota na lista.", vbExclamation: Exit Sub
    If MsgBox("Excluir a alíquota de " & txtAliq.Text & "% (início " & txtIni.Text & ")?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrD
    vgDb.Execute "DELETE FROM CBS_Aliquota WHERE Id=" & iId
    AplicarAliquotasReforma
    MsgBox "Excluído.", vbInformation
    Carregar
    Exit Sub
ErrD:
    MsgBox "Erro ao excluir: " & Err.Description, vbCritical
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
