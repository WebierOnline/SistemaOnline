VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIBSCBS_Cadastro
   Caption         =   "Cadastro IBS/CBS — CST e Classificação Tributária"
   ClientHeight    =   9840
   ClientLeft      =   120
   ClientTop       =   432
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid lstCST
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   5292
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid lstClassTrib
      Height          =   2400
      Left            =   5160
      TabIndex        =   1
      Top             =   480
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   4233
      _Version        =   393216
   End
   Begin VB.Label lblCSTHeader
      Caption         =   "CST IBS/CBS"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4800
      FontBold        =   -1  'True
   End
   Begin VB.Label lblCST
      Caption         =   "CST:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   480
   End
   Begin VB.TextBox txtCST
      Height          =   288
      Left            =   720
      MaxLength       =   3
      TabIndex        =   4
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label lblDescCST
      Caption         =   "Descrição:"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   960
   End
   Begin VB.TextBox txtDescCST
      Height          =   288
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   4800
   End
   Begin VB.Frame frachkCST
      Caption         =   "Indicadores"
      Height          =   1200
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   4800
      Begin VB.CheckBox chkCST
         Caption         =   "gIBSCBS"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "gIBSCBSMono"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "gRed"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "gDif"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   11
         Top             =   450
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "gTransfCred"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "gCredPresIBSZFM"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   13
         Top             =   660
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "gAjusteCompet"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   870
         Width           =   2280
      End
      Begin VB.CheckBox chkCST
         Caption         =   "RedutorBC"
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   15
         Top             =   870
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdNovoCST
      Caption         =   "Novo CST"
      Height          =   360
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   1400
   End
   Begin VB.CommandButton cmdSalvarCST
      Caption         =   "Salvar CST"
      Height          =   360
      Left            =   1680
      TabIndex        =   17
      Top             =   5880
      Width           =   1440
   End
   Begin VB.CommandButton cmdExcluirCST
      Caption         =   "Excluir CST"
      Height          =   360
      Left            =   3240
      TabIndex        =   18
      Top             =   5880
      Width           =   1680
   End
   Begin VB.Label lblClsHeader
      Caption         =   "Classificação Tributária (CST selecionado)"
      Height          =   240
      Left            =   5160
      TabIndex        =   19
      Top             =   120
      Width           =   9120
      FontBold        =   -1  'True
   End
   Begin VB.Label lblCClassTrib
      Caption         =   "Cód. Classe:"
      Height          =   240
      Left            =   5160
      TabIndex        =   20
      Top             =   3000
      Width           =   1440
   End
   Begin VB.TextBox txtCClassTrib
      Height          =   288
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   21
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label lblNomeCls
      Caption         =   "Nome:"
      Height          =   240
      Left            =   7800
      TabIndex        =   22
      Top             =   3000
      Width           =   600
   End
   Begin VB.TextBox txtNomeCls
      Height          =   288
      Left            =   8520
      TabIndex        =   23
      Top             =   3000
      Width           =   5760
   End
   Begin VB.Label lblDescCls
      Caption         =   "Descrição:"
      Height          =   240
      Left            =   5160
      TabIndex        =   24
      Top             =   3360
      Width           =   1440
   End
   Begin VB.TextBox txtDescricaoCls
      Height          =   600
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   3600
      Width           =   9120
   End
   Begin VB.Label lblLCRedacao
      Caption         =   "LC Redação:"
      Height          =   240
      Left            =   5160
      TabIndex        =   26
      Top             =   4260
      Width           =   1440
   End
   Begin VB.TextBox txtLCRedacao
      Height          =   600
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   4500
      Width           =   9120
   End
   Begin VB.Label lblLC21425
      Caption         =   "LC 214/25:"
      Height          =   240
      Left            =   5160
      TabIndex        =   28
      Top             =   5160
      Width           =   1200
   End
   Begin VB.TextBox txtLC21425
      Height          =   288
      Left            =   6480
      TabIndex        =   29
      Top             =   5160
      Width           =   2400
   End
   Begin VB.Label lblTipoAliq
      Caption         =   "Tipo Alíquota:"
      Height          =   240
      Left            =   9000
      TabIndex        =   30
      Top             =   5160
      Width           =   1560
   End
   Begin VB.TextBox txtTipoAliq
      Height          =   288
      Left            =   10680
      TabIndex        =   31
      Top             =   5160
      Width           =   2640
   End
   Begin VB.Label lblPRedIBS
      Caption         =   "% Red. IBS:"
      Height          =   240
      Left            =   5160
      TabIndex        =   32
      Top             =   5520
      Width           =   1200
   End
   Begin VB.TextBox txtPRedIBS
      Height          =   288
      Left            =   6480
      TabIndex        =   33
      Top             =   5520
      Width           =   960
   End
   Begin VB.Label lblPRedCBS
      Caption         =   "% Red. CBS:"
      Height          =   240
      Left            =   7560
      TabIndex        =   34
      Top             =   5520
      Width           =   1200
   End
   Begin VB.TextBox txtPRedCBS
      Height          =   288
      Left            =   8880
      TabIndex        =   35
      Top             =   5520
      Width           =   960
   End
   Begin VB.Label lblCreditoPara
      Caption         =   "Crédito para:"
      Height          =   240
      Left            =   9960
      TabIndex        =   36
      Top             =   5520
      Width           =   1440
   End
   Begin VB.TextBox txtCreditoPara
      Height          =   288
      Left            =   11520
      TabIndex        =   37
      Top             =   5520
      Width           =   2760
   End
   Begin VB.Label lblDIniVig
      Caption         =   "Vig. Ini:"
      Height          =   240
      Left            =   5160
      TabIndex        =   38
      Top             =   5880
      Width           =   1080
   End
   Begin VB.TextBox txtDIniVig
      Height          =   288
      Left            =   6360
      TabIndex        =   39
      Top             =   5880
      Width           =   1440
   End
   Begin VB.Label lblDFimVig
      Caption         =   "Vig. Fim:"
      Height          =   240
      Left            =   7920
      TabIndex        =   40
      Top             =   5880
      Width           =   1080
   End
   Begin VB.TextBox txtDFimVig
      Height          =   288
      Left            =   9120
      TabIndex        =   41
      Top             =   5880
      Width           =   1440
   End
   Begin VB.Label lblAnexo
      Caption         =   "Anexo:"
      Height          =   240
      Left            =   10680
      TabIndex        =   42
      Top             =   5880
      Width           =   720
   End
   Begin VB.TextBox txtAnexo
      Height          =   288
      Left            =   11520
      TabIndex        =   43
      Top             =   5880
      Width           =   720
   End
   Begin VB.Label lblLink
      Caption         =   "Link:"
      Height          =   240
      Left            =   5160
      TabIndex        =   44
      Top             =   6240
      Width           =   600
   End
   Begin VB.TextBox txtLink
      Height          =   288
      Left            =   5880
      TabIndex        =   45
      Top             =   6240
      Width           =   8400
   End
   Begin VB.Frame frachkCls
      Caption         =   "Indicadores NF-e"
      Height          =   1920
      Left            =   5160
      TabIndex        =   46
      Top             =   6600
      Width           =   9120
      Begin VB.CheckBox chkCls
         Caption         =   "gTribRegular"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   300
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "gCredPresOper"
         Height          =   210
         Index           =   1
         Left            =   2400
         TabIndex        =   48
         Top             =   300
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "gMonoPadrao"
         Height          =   210
         Index           =   2
         Left            =   4680
         TabIndex        =   49
         Top             =   300
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indMonoReten"
         Height          =   210
         Index           =   3
         Left            =   6960
         TabIndex        =   50
         Top             =   300
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indMonoRet"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   51
         Top             =   540
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indMonoDif"
         Height          =   210
         Index           =   5
         Left            =   2400
         TabIndex        =   52
         Top             =   540
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "gEstornoCred"
         Height          =   210
         Index           =   6
         Left            =   4680
         TabIndex        =   53
         Top             =   540
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFeABI"
         Height          =   210
         Index           =   7
         Left            =   6960
         TabIndex        =   54
         Top             =   540
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFe"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   55
         Top             =   780
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFCe"
         Height          =   210
         Index           =   9
         Left            =   2400
         TabIndex        =   56
         Top             =   780
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indCTe"
         Height          =   210
         Index           =   10
         Left            =   4680
         TabIndex        =   57
         Top             =   780
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indCTeOS"
         Height          =   210
         Index           =   11
         Left            =   6960
         TabIndex        =   58
         Top             =   780
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indBPe"
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   59
         Top             =   1020
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indBPeTA"
         Height          =   210
         Index           =   13
         Left            =   2400
         TabIndex        =   60
         Top             =   1020
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indBPeTM"
         Height          =   210
         Index           =   14
         Left            =   4680
         TabIndex        =   61
         Top             =   1020
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNF3e"
         Height          =   210
         Index           =   15
         Left            =   6960
         TabIndex        =   62
         Top             =   1020
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFSe"
         Height          =   210
         Index           =   16
         Left            =   120
         TabIndex        =   63
         Top             =   1260
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFSe_Via"
         Height          =   210
         Index           =   17
         Left            =   2400
         TabIndex        =   64
         Top             =   1260
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFCom"
         Height          =   210
         Index           =   18
         Left            =   4680
         TabIndex        =   65
         Top             =   1260
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFAg"
         Height          =   210
         Index           =   19
         Left            =   6960
         TabIndex        =   66
         Top             =   1260
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indNFGas"
         Height          =   210
         Index           =   20
         Left            =   120
         TabIndex        =   67
         Top             =   1500
         Width           =   2160
      End
      Begin VB.CheckBox chkCls
         Caption         =   "indDERE"
         Height          =   210
         Index           =   21
         Left            =   2400
         TabIndex        =   68
         Top             =   1500
         Width           =   2160
      End
   End
   Begin VB.CommandButton cmdNovaClasse
      Caption         =   "Nova Classe"
      Height          =   360
      Left            =   5160
      TabIndex        =   69
      Top             =   8640
      Width           =   1800
   End
   Begin VB.CommandButton cmdSalvarClasse
      Caption         =   "Salvar Classe"
      Height          =   360
      Left            =   7200
      TabIndex        =   70
      Top             =   8640
      Width           =   1800
   End
   Begin VB.CommandButton cmdExcluirClasse
      Caption         =   "Excluir Classe"
      Height          =   360
      Left            =   9240
      TabIndex        =   71
      Top             =   8640
      Width           =   1920
   End
   Begin VB.Line linSep
      X1              =   5040
      X2              =   5040
      Y1              =   120
      Y2              =   9600
   End
End
Attribute VB_Name = "frmIBSCBS_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bNovoCST    As Boolean
Dim bNovaClasse As Boolean
Dim sCSTSel     As String
Dim sClsSel     As String

Private Sub Form_Load()
    With lstCST
        .Cols = 2
        .ColWidth(0) = 900
        .ColWidth(1) = 3780
        .Row = 0: .Col = 0: .Text = "CST"
        .Row = 0: .Col = 1: .Text = "Descri" & Chr(231) & Chr(227) & "o"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
    With lstClassTrib
        .Cols = 3
        .ColWidth(0) = 900
        .ColWidth(1) = 3600
        .ColWidth(2) = 4500
        .Row = 0: .Col = 0: .Text = "Classe"
        .Row = 0: .Col = 1: .Text = "Nome"
        .Row = 0: .Col = 2: .Text = "Descri" & Chr(231) & Chr(227) & "o"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
    CarregarCSTs
End Sub

Private Sub CarregarCSTs()
    Dim rRs As ADODB.Recordset
    Dim r   As Long
    lstCST.Rows = 1
    RsOpen rRs, "SELECT CST, DescricaoIBSCBS FROM TbIBSCBS ORDER BY CST"
    Do While Not rRs.EOF
        lstCST.AddItem ""
        r = lstCST.Rows - 1
        lstCST.Row = r: lstCST.Col = 0: lstCST.Text = rRs("CST")
        lstCST.Row = r: lstCST.Col = 1: lstCST.Text = rRs("DescricaoIBSCBS")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub lstCST_Click()
    If lstCST.Row = 0 Then Exit Sub
    lstCST.Col = 0
    sCSTSel = lstCST.Text
    CarregarDadosCST sCSTSel
    CarregarClassTribs sCSTSel
    bNovoCST = False
End Sub

Private Sub CarregarDadosCST(ByVal sCST As String)
    Dim rRs As ADODB.Recordset
    RsOpen rRs, "SELECT * FROM TbIBSCBS WHERE CST = '" & sCST & "'"
    If rRs.EOF Then If rRs.State <> 0 Then rRs.Close: Exit Sub
    txtCST.Text    = rRs("CST")
    txtDescCST.Text = rRs("DescricaoIBSCBS")
    chkCST(0).Value = IIf(rRs("ind_gIBSCBS") <> 0, 1, 0)
    chkCST(1).Value = IIf(rRs("ind_gIBSCBSMono") <> 0, 1, 0)
    chkCST(2).Value = IIf(rRs("ind_gRed") <> 0, 1, 0)
    chkCST(3).Value = IIf(rRs("ind_gDif") <> 0, 1, 0)
    chkCST(4).Value = IIf(rRs("ind_gTransfCred") <> 0, 1, 0)
    chkCST(5).Value = IIf(rRs("ind_gCredPresIBSZFM") <> 0, 1, 0)
    chkCST(6).Value = IIf(rRs("ind_gAjusteCompet") <> 0, 1, 0)
    chkCST(7).Value = IIf(rRs("ind_RedutorBC") <> 0, 1, 0)
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub CarregarClassTribs(ByVal sCST As String)
    Dim rRs As ADODB.Recordset
    Dim r   As Long
    lstClassTrib.Rows = 1
    RsOpen rRs, "SELECT cClassTrib, NomecClassTrib, DescricaocClassTrib FROM TbIBSCBSClassTrib WHERE CST = '" & sCST & "' ORDER BY cClassTrib"
    Do While Not rRs.EOF
        lstClassTrib.AddItem ""
        r = lstClassTrib.Rows - 1
        lstClassTrib.Row = r: lstClassTrib.Col = 0: lstClassTrib.Text = rRs("cClassTrib")
        lstClassTrib.Row = r: lstClassTrib.Col = 1: lstClassTrib.Text = rRs("NomecClassTrib")
        lstClassTrib.Row = r: lstClassTrib.Col = 2: lstClassTrib.Text = rRs("DescricaocClassTrib")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    LimparCamposClasse
    sClsSel = ""
    bNovaClasse = False
End Sub

Private Sub lstClassTrib_Click()
    If lstClassTrib.Row = 0 Then Exit Sub
    lstClassTrib.Col = 0
    sClsSel = lstClassTrib.Text
    CarregarDadosClasse sCSTSel, sClsSel
    bNovaClasse = False
End Sub

Private Sub CarregarDadosClasse(ByVal sCST As String, ByVal sCls As String)
    Dim rRs As ADODB.Recordset
    Dim i   As Integer
    RsOpen rRs, "SELECT * FROM TbIBSCBSClassTrib WHERE CST = '" & sCST & "' AND cClassTrib = '" & sCls & "'"
    If rRs.EOF Then If rRs.State <> 0 Then rRs.Close: Exit Sub
    txtCClassTrib.Text   = rRs("cClassTrib")
    txtNomeCls.Text      = rRs("NomecClassTrib")
    txtDescricaoCls.Text = rRs("DescricaocClassTrib")
    txtLCRedacao.Text    = IIf(IsNull(rRs("LC_Redacao")),    "", rRs("LC_Redacao"))
    txtLC21425.Text      = IIf(IsNull(rRs("LC_214_25")),     "", rRs("LC_214_25"))
    txtTipoAliq.Text     = IIf(IsNull(rRs("TipoDeAliquota")),"", rRs("TipoDeAliquota"))
    txtPRedIBS.Text      = IIf(IsNull(rRs("pRedIBS")),  "0", rRs("pRedIBS"))
    txtPRedCBS.Text      = IIf(IsNull(rRs("pRedCBS")),  "0", rRs("pRedCBS"))
    txtCreditoPara.Text  = IIf(IsNull(rRs("Credito_para")),  "", rRs("Credito_para"))
    txtDIniVig.Text      = IIf(IsNull(rRs("dIniVig")), "", Format(rRs("dIniVig"), "yyyy-mm-dd"))
    txtDFimVig.Text      = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "yyyy-mm-dd"))
    txtAnexo.Text        = IIf(IsNull(rRs("Anexo")), "", rRs("Anexo"))
    txtLink.Text         = IIf(IsNull(rRs("Link")),  "", rRs("Link"))
    Dim campos(21) As String
    campos(0)="ind_gTribRegular": campos(1)="ind_gCredPresOper": campos(2)="ind_gMonoPadrao"
    campos(3)="indMonoReten":     campos(4)="indMonoRet":        campos(5)="indMonoDif"
    campos(6)="ind_gEstornoCred": campos(7)="indNFeABI":         campos(8)="indNFe"
    campos(9)="indNFCe":          campos(10)="indCTe":           campos(11)="indCTeOS"
    campos(12)="indBPe":          campos(13)="indBPeTA":         campos(14)="indBPeTM"
    campos(15)="indNF3e":         campos(16)="indNFSe":          campos(17)="indNFSe_Via"
    campos(18)="indNFCom":        campos(19)="indNFAg":          campos(20)="indNFGas"
    campos(21)="indDERE"
    For i = 0 To 21
        chkCls(i).Value = IIf(rRs(campos(i)) <> 0, 1, 0)
    Next i
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub LimparCamposClasse()
    Dim i As Integer
    txtCClassTrib.Text = "": txtNomeCls.Text = "": txtDescricaoCls.Text = ""
    txtLCRedacao.Text  = "": txtLC21425.Text = "": txtTipoAliq.Text     = ""
    txtPRedIBS.Text    = "0": txtPRedCBS.Text = "0": txtCreditoPara.Text = ""
    txtDIniVig.Text    = "": txtDFimVig.Text = "": txtAnexo.Text = "": txtLink.Text = ""
    For i = 0 To 21: chkCls(i).Value = 0: Next i
End Sub

Private Sub LimparCamposCST()
    Dim i As Integer
    txtCST.Text = "": txtDescCST.Text = ""
    For i = 0 To 7: chkCST(i).Value = 0: Next i
End Sub

' -- CST CRUD ----------------------------------------------------------------
Private Sub cmdNovoCST_Click()
    bNovoCST = True
    LimparCamposCST
    txtCST.SetFocus
End Sub

Private Sub cmdSalvarCST_Click()
    Dim sCST As String, sDesc As String, sql As String
    Dim ind(7) As Integer, i As Integer
    sCST = Trim(txtCST.Text): sDesc = Trim(txtDescCST.Text)
    If sCST = "" Or sDesc = "" Then MsgBox "CST e Descri" & Chr(231) & Chr(227) & "o s" & Chr(227) & "o obrigat" & Chr(243) & "rios.", vbExclamation: Exit Sub
    For i = 0 To 7: ind(i) = IIf(chkCST(i).Value = 1, 1, 0): Next i
    If bNovoCST Then
        sql = "INSERT INTO TbIBSCBS (CST,DescricaoIBSCBS,ind_gIBSCBS,ind_gIBSCBSMono,ind_gRed,ind_gDif,ind_gTransfCred,ind_gCredPresIBSZFM,ind_gAjusteCompet,ind_RedutorBC) VALUES ('" & sCST & "','" & Replace(sDesc, "'", "''") & "'," & ind(0) & "," & ind(1) & "," & ind(2) & "," & ind(3) & "," & ind(4) & "," & ind(5) & "," & ind(6) & "," & ind(7) & ")"
    Else
        sql = "UPDATE TbIBSCBS SET DescricaoIBSCBS='" & Replace(sDesc, "'", "''") & "',ind_gIBSCBS=" & ind(0) & ",ind_gIBSCBSMono=" & ind(1) & ",ind_gRed=" & ind(2) & ",ind_gDif=" & ind(3) & ",ind_gTransfCred=" & ind(4) & ",ind_gCredPresIBSZFM=" & ind(5) & ",ind_gAjusteCompet=" & ind(6) & ",ind_RedutorBC=" & ind(7) & " WHERE CST='" & sCSTSel & "'"
    End If
    On Error GoTo ErrCST
    vgDb.Execute sql
    MsgBox "CST salvo.", vbInformation
    sCSTSel = sCST: bNovoCST = False: CarregarCSTs
    Exit Sub
ErrCST: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluirCST_Click()
    If sCSTSel = "" Then Exit Sub
    If MsgBox("Excluir CST " & sCSTSel & " e TODAS as suas classifica" & Chr(231) & Chr(245) & "es?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrDelCST
    vgDb.Execute "DELETE FROM TbIBSCBSClassTrib WHERE CST='" & sCSTSel & "'"
    vgDb.Execute "DELETE FROM TbIBSCBS WHERE CST='" & sCSTSel & "'"
    MsgBox "CST exclu" & Chr(237) & "do.", vbInformation
    sCSTSel = "": LimparCamposCST: lstClassTrib.Rows = 1: LimparCamposClasse: CarregarCSTs
    Exit Sub
ErrDelCST: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' -- CLASSTRIB CRUD -----------------------------------------------------------
Private Sub cmdNovaClasse_Click()
    If sCSTSel = "" Then MsgBox "Selecione um CST primeiro.", vbExclamation: Exit Sub
    bNovaClasse = True: LimparCamposClasse
    txtDIniVig.Text = Format(Now, "yyyy-mm-dd"): txtCClassTrib.SetFocus
End Sub

Private Function IndCls() As String
    Dim i As Integer, s As String
    Dim campos(21) As String
    campos(0)="ind_gTribRegular": campos(1)="ind_gCredPresOper": campos(2)="ind_gMonoPadrao"
    campos(3)="indMonoReten":     campos(4)="indMonoRet":        campos(5)="indMonoDif"
    campos(6)="ind_gEstornoCred": campos(7)="indNFeABI":         campos(8)="indNFe"
    campos(9)="indNFCe":          campos(10)="indCTe":           campos(11)="indCTeOS"
    campos(12)="indBPe":          campos(13)="indBPeTA":         campos(14)="indBPeTM"
    campos(15)="indNF3e":         campos(16)="indNFSe":          campos(17)="indNFSe_Via"
    campos(18)="indNFCom":        campos(19)="indNFAg":          campos(20)="indNFGas"
    campos(21)="indDERE"
    For i = 0 To 21
        s = s & campos(i) & "=" & IIf(chkCls(i).Value = 1, 1, 0) & IIf(i < 21, ",", "")
    Next i
    IndCls = s
End Function

Private Sub cmdSalvarClasse_Click()
    Dim sCls As String, sql As String
    Dim sDIni As String, sDFim As String
    If sCSTSel = "" Then MsgBox "Selecione um CST.", vbExclamation: Exit Sub
    sCls = Trim(txtCClassTrib.Text)
    If sCls = "" Or Trim(txtNomeCls.Text) = "" Then MsgBox "C" & Chr(243) & "d. Classe e Nome s" & Chr(227) & "o obrigat" & Chr(243) & "rios.", vbExclamation: Exit Sub
    sDIni = IIf(Trim(txtDIniVig.Text) = "", "NULL", "CONVERT(date,'" & Trim(txtDIniVig.Text) & "',23)")
    sDFim = IIf(Trim(txtDFimVig.Text) = "", "NULL", "CONVERT(date,'" & Trim(txtDFimVig.Text) & "',23)")
    If bNovaClasse Then
        sql = "INSERT INTO TbIBSCBSClassTrib (CST,DescricaoIBSCBS,cClassTrib,NomecClassTrib,DescricaocClassTrib,LC_Redacao,LC_214_25,TipoDeAliquota,pRedIBS,pRedCBS,Credito_para,dIniVig,dFimVig,Anexo,Link," & IndCls() & ") VALUES ('" & sCSTSel & "',(SELECT DescricaoIBSCBS FROM TbIBSCBS WHERE CST='" & sCSTSel & "'),'" & sCls & "','" & Replace(Trim(txtNomeCls.Text), "'", "''") & "','" & Replace(Trim(txtDescricaoCls.Text), "'", "''") & "','" & Replace(Trim(txtLCRedacao.Text), "'", "''") & "','" & Replace(Trim(txtLC21425.Text), "'", "''") & "','" & Replace(Trim(txtTipoAliq.Text), "'", "''") & "'," & Val(txtPRedIBS.Text) & "," & Val(txtPRedCBS.Text) & ",'" & Replace(Trim(txtCreditoPara.Text), "'", "''") & "'," & sDIni & "," & sDFim & ",'" & Replace(Trim(txtAnexo.Text), "'", "''") & "','" & Replace(Trim(txtLink.Text), "'", "''") & "'"
        ' close VALUES - replace campo=valor format for INSERT
        ' rebuild properly:
        Dim i As Integer, sCampos As String, sVals As String
        Dim campos(21) As String, ind(21) As Integer
        campos(0)="ind_gTribRegular": campos(1)="ind_gCredPresOper": campos(2)="ind_gMonoPadrao"
        campos(3)="indMonoReten":     campos(4)="indMonoRet":        campos(5)="indMonoDif"
        campos(6)="ind_gEstornoCred": campos(7)="indNFeABI":         campos(8)="indNFe"
        campos(9)="indNFCe":          campos(10)="indCTe":           campos(11)="indCTeOS"
        campos(12)="indBPe":          campos(13)="indBPeTA":         campos(14)="indBPeTM"
        campos(15)="indNF3e":         campos(16)="indNFSe":          campos(17)="indNFSe_Via"
        campos(18)="indNFCom":        campos(19)="indNFAg":          campos(20)="indNFGas"
        campos(21)="indDERE"
        For i = 0 To 21: ind(i) = IIf(chkCls(i).Value = 1, 1, 0): sCampos = sCampos & "," & campos(i): sVals = sVals & "," & ind(i): Next i
        sql = "INSERT INTO TbIBSCBSClassTrib (CST,DescricaoIBSCBS,cClassTrib,NomecClassTrib,DescricaocClassTrib,LC_Redacao,LC_214_25,TipoDeAliquota,pRedIBS,pRedCBS,Credito_para,dIniVig,dFimVig,Anexo,Link" & sCampos & ") VALUES ('" & sCSTSel & "',(SELECT DescricaoIBSCBS FROM TbIBSCBS WHERE CST='" & sCSTSel & "'),'" & sCls & "','" & Replace(Trim(txtNomeCls.Text), "'", "''") & "','" & Replace(Trim(txtDescricaoCls.Text), "'", "''") & "','" & Replace(Trim(txtLCRedacao.Text), "'", "''") & "','" & Replace(Trim(txtLC21425.Text), "'", "''") & "','" & Replace(Trim(txtTipoAliq.Text), "'", "''") & "'," & Val(txtPRedIBS.Text) & "," & Val(txtPRedCBS.Text) & ",'" & Replace(Trim(txtCreditoPara.Text), "'", "''") & "'," & sDIni & "," & sDFim & ",'" & Replace(Trim(txtAnexo.Text), "'", "''") & "','" & Replace(Trim(txtLink.Text), "'", "''") & "'" & sVals & ")"
    Else
        sql = "UPDATE TbIBSCBSClassTrib SET NomecClassTrib='" & Replace(Trim(txtNomeCls.Text), "'", "''") & "',DescricaocClassTrib='" & Replace(Trim(txtDescricaoCls.Text), "'", "''") & "',LC_Redacao='" & Replace(Trim(txtLCRedacao.Text), "'", "''") & "',LC_214_25='" & Replace(Trim(txtLC21425.Text), "'", "''") & "',TipoDeAliquota='" & Replace(Trim(txtTipoAliq.Text), "'", "''") & "',pRedIBS=" & Val(txtPRedIBS.Text) & ",pRedCBS=" & Val(txtPRedCBS.Text) & ",Credito_para='" & Replace(Trim(txtCreditoPara.Text), "'", "''") & "',dIniVig=" & sDIni & ",dFimVig=" & sDFim & ",Anexo='" & Replace(Trim(txtAnexo.Text), "'", "''") & "',Link='" & Replace(Trim(txtLink.Text), "'", "''") & "'," & IndCls() & " WHERE CST='" & sCSTSel & "' AND cClassTrib='" & sClsSel & "'"
    End If
    On Error GoTo ErrCls
    vgDb.Execute sql
    MsgBox "Classifica" & Chr(231) & Chr(227) & "o salva.", vbInformation
    sClsSel = sCls: bNovaClasse = False: CarregarClassTribs sCSTSel
    Exit Sub
ErrCls: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluirClasse_Click()
    If sClsSel = "" Then Exit Sub
    If MsgBox("Excluir classifica" & Chr(231) & Chr(227) & "o " & sClsSel & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrDel
    vgDb.Execute "DELETE FROM TbIBSCBSClassTrib WHERE CST='" & sCSTSel & "' AND cClassTrib='" & sClsSel & "'"
    MsgBox "Exclu" & Chr(237) & "do.", vbInformation
    sClsSel = "": LimparCamposClasse: CarregarClassTribs sCSTSel
    Exit Sub
ErrDel: MsgBox "Erro: " & Err.Description, vbCritical
End Sub
