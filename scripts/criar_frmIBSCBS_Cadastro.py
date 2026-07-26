# -*- coding: utf-8 -*-
path = 'C:/Projeto/OnlineCommerce/Forms/frmIBSCBS_Cadastro.frm'

ti = [0]
def T():
    v = ti[0]; ti[0] += 1; return v

cst_chk_caps  = ['gIBSCBS','gIBSCBSMono','gRed','gDif','gTransfCred','gCredPresIBSZFM','gAjusteCompet','RedutorBC']
cls_chk_caps  = ['gTribRegular','gCredPresOper','gMonoPadrao','indMonoReten','indMonoRet','indMonoDif',
                 'gEstornoCred','indNFeABI','indNFe','indNFCe','indCTe','indCTeOS',
                 'indBPe','indBPeTA','indBPeTM','indNF3e','indNFSe','indNFSe_Via',
                 'indNFCom','indNFAg','indNFGas','indDERE']

def lbl(name, cap, left, top, width, height=240, bold=False):
    b = '   True' if bold else '   False'
    ti_v = T()
    return (
        f'   Begin VB.Label {name}\n'
        f'      Caption         =   "{cap}"\n'
        f'      Height          =   {height}\n'
        f'      Left            =   {left}\n'
        f'      TabIndex        =   {ti_v}\n'
        f'      Top             =   {top}\n'
        f'      Width           =   {width}\n'
        + (f'      FontBold        =   -1  \'True\n' if bold else '')
        + f'   End\n'
    )

def txt(name, left, top, width, height=288, multiline=False, maxlen=0):
    ti_v = T()
    s  = f'   Begin VB.TextBox {name}\n'
    s += f'      Height          =   {height}\n'
    s += f'      Left            =   {left}\n'
    if maxlen: s += f'      MaxLength       =   {maxlen}\n'
    if multiline:
        s += f'      MultiLine       =   -1  \'True\n'
        s += f'      ScrollBars      =   2  \'Vertical\n'
    s += f'      TabIndex        =   {ti_v}\n'
    s += f'      Top             =   {top}\n'
    s += f'      Width           =   {width}\n'
    s += f'   End\n'
    return s

def cmd(name, cap, left, top, width=1680, height=360):
    ti_v = T()
    return (
        f'   Begin VB.CommandButton {name}\n'
        f'      Caption         =   "{cap}"\n'
        f'      Height          =   {height}\n'
        f'      Left            =   {left}\n'
        f'      TabIndex        =   {ti_v}\n'
        f'      Top             =   {top}\n'
        f'      Width           =   {width}\n'
        f'   End\n'
    )

def grid(name, left, top, width, height):
    ti_v = T()
    ex = round(width * 2540 / 1440)
    ey = round(height * 2540 / 1440)
    return (
        f'   Begin MSFlexGridLib.MSFlexGrid {name}\n'
        f'      Height          =   {height}\n'
        f'      Left            =   {left}\n'
        f'      TabIndex        =   {ti_v}\n'
        f'      Top             =   {top}\n'
        f'      Width           =   {width}\n'
        f'      _ExtentX        =   {ex}\n'
        f'      _ExtentY        =   {ey}\n'
        f'      _Version        =   393216\n'
        f'   End\n'
    )

def chk_array(arr_name, caps, left_frame, top_frame, frame_w, frame_h, frame_cap,
              col_starts, row_tops, chk_w, chk_h, indent='   '):
    i2 = indent + '   '
    s  = f'{indent}Begin VB.Frame fra{arr_name}\n'
    s += f'{i2}Caption         =   "{frame_cap}"\n'
    s += f'{i2}Height          =   {frame_h}\n'
    s += f'{i2}Left            =   {left_frame}\n'
    s += f'{i2}TabIndex        =   {T()}\n'
    s += f'{i2}Top             =   {top_frame}\n'
    s += f'{i2}Width           =   {frame_w}\n'
    ncols = len(col_starts)
    nrows = len(row_tops)
    for idx, cap in enumerate(caps):
        col = idx % ncols
        row = idx // ncols
        if row >= nrows: break
        s += f'{i2}Begin VB.CheckBox {arr_name}\n'
        s += f'{i2}   Caption         =   "{cap}"\n'
        s += f'{i2}   Height          =   {chk_h}\n'
        s += f'{i2}   Index           =   {idx}\n'
        s += f'{i2}   Left            =   {col_starts[col]}\n'
        s += f'{i2}   TabIndex        =   {T()}\n'
        s += f'{i2}   Top             =   {row_tops[row]}\n'
        s += f'{i2}   Width           =   {chk_w}\n'
        s += f'{i2}End\n'
    s += f'{indent}End\n'
    return s

# -- FORM STRUCTURE ------------------------------------------------------------
frm_def  = 'VERSION 5.00\n'
frm_def += 'Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSFLXGRD.OCX"\n'
frm_def += 'Begin VB.Form frmIBSCBS_Cadastro\n'
frm_def += '   Caption         =   "Cadastro IBS/CBS \u2014 CST e Classifica\xe7\xe3o Tribut\xe1ria"\n'
frm_def += '   ClientHeight    =   9840\n'
frm_def += '   ClientLeft      =   120\n'
frm_def += '   ClientTop       =   432\n'
frm_def += '   ClientWidth     =   14400\n'
frm_def += '   LinkTopic       =   "Form1"\n'
frm_def += '   ScaleHeight     =   9840\n'
frm_def += '   ScaleWidth      =   14400\n'
frm_def += '   StartUpPosition =   2  \'CenterScreen\n'

# Grids
frm_def += grid('lstCST',       120,  480, 4800, 3000)
frm_def += grid('lstClassTrib', 5160, 480, 9120, 2400)

# CST edit fields
frm_def += lbl('lblCSTHeader', 'CST IBS/CBS', 120, 120, 4800, bold=True)
frm_def += lbl('lblCST',       'CST:',         120, 3600, 480)
frm_def += txt('txtCST',                        720, 3600, 720, maxlen=3)
frm_def += lbl('lblDescCST',   'Descri\xe7\xe3o:', 120, 3960, 960)
frm_def += txt('txtDescCST',                    120, 4200, 4800)

# CST indicators frame
frm_def += chk_array(
    'chkCST', cst_chk_caps,
    left_frame=120, top_frame=4560, frame_w=4800, frame_h=1200,
    frame_cap='Indicadores',
    col_starts=[120, 2520], row_tops=[240, 450, 660, 870],
    chk_w=2280, chk_h=195
)

# CST buttons
frm_def += cmd('cmdNovoCST',    'Novo CST',    120,  5880, 1400)
frm_def += cmd('cmdSalvarCST',  'Salvar CST',  1680, 5880, 1440)
frm_def += cmd('cmdExcluirCST', 'Excluir CST', 3240, 5880, 1680)

# ClassTrib header + edit fields
frm_def += lbl('lblClsHeader', 'Classifica\xe7\xe3o Tribut\xe1ria (CST selecionado)', 5160, 120, 9120, bold=True)
frm_def += lbl('lblCClassTrib','C\xf3d. Classe:',  5160, 3000, 1440)
frm_def += txt('txtCClassTrib',                     6720, 3000, 960,  maxlen=6)
frm_def += lbl('lblNomeCls',   'Nome:',             7800, 3000, 600)
frm_def += txt('txtNomeCls',                        8520, 3000, 5760)
frm_def += lbl('lblDescCls',   'Descri\xe7\xe3o:', 5160, 3360, 1440)
frm_def += txt('txtDescricaoCls',                   5160, 3600, 9120, height=600, multiline=True)
frm_def += lbl('lblLCRedacao', 'LC Reda\xe7\xe3o:',5160, 4260, 1440)
frm_def += txt('txtLCRedacao',                      5160, 4500, 9120, height=600, multiline=True)
frm_def += lbl('lblLC21425',   'LC 214/25:',        5160, 5160, 1200)
frm_def += txt('txtLC21425',                        6480, 5160, 2400)
frm_def += lbl('lblTipoAliq',  'Tipo Al\xedquota:', 9000, 5160, 1560)
frm_def += txt('txtTipoAliq',                       10680,5160, 2640)
frm_def += lbl('lblPRedIBS',   '% Red. IBS:',       5160, 5520, 1200)
frm_def += txt('txtPRedIBS',                        6480, 5520, 960)
frm_def += lbl('lblPRedCBS',   '% Red. CBS:',       7560, 5520, 1200)
frm_def += txt('txtPRedCBS',                        8880, 5520, 960)
frm_def += lbl('lblCreditoPara','Cr\xe9dito para:', 9960, 5520, 1440)
frm_def += txt('txtCreditoPara',                    11520,5520, 2760)
frm_def += lbl('lblDIniVig',   'Vig. Ini:',         5160, 5880, 1080)
frm_def += txt('txtDIniVig',                        6360, 5880, 1440)
frm_def += lbl('lblDFimVig',   'Vig. Fim:',         7920, 5880, 1080)
frm_def += txt('txtDFimVig',                        9120, 5880, 1440)
frm_def += lbl('lblAnexo',     'Anexo:',            10680,5880, 720)
frm_def += txt('txtAnexo',                          11520,5880, 720)
frm_def += lbl('lblLink',      'Link:',             5160, 6240, 600)
frm_def += txt('txtLink',                           5880, 6240, 8400)

# ClassTrib indicators frame (22 chk, 4 cols x 6 rows)
col_starts_cls = [120, 2400, 4680, 6960]
row_tops_cls   = [300, 540,  780,  1020, 1260, 1500]
frm_def += chk_array(
    'chkCls', cls_chk_caps,
    left_frame=5160, top_frame=6600, frame_w=9120, frame_h=1920,
    frame_cap='Indicadores NF-e',
    col_starts=col_starts_cls, row_tops=row_tops_cls,
    chk_w=2160, chk_h=210
)

# ClassTrib buttons
frm_def += cmd('cmdNovaClasse',    'Nova Classe',    5160, 8640, 1800)
frm_def += cmd('cmdSalvarClasse',  'Salvar Classe',  7200, 8640, 1800)
frm_def += cmd('cmdExcluirClasse', 'Excluir Classe', 9240, 8640, 1920)

# Vertical separator
frm_def += (
    '   Begin VB.Line linSep\n'
    '      X1              =   5040\n'
    '      X2              =   5040\n'
    '      Y1              =   120\n'
    '      Y2              =   9600\n'
    '   End\n'
)

frm_def += 'End\n'

# -- ATTRIBUTES ---------------------------------------------------------------
frm_def += 'Attribute VB_Name = "frmIBSCBS_Cadastro"\n'
frm_def += 'Attribute VB_GlobalNameSpace = False\n'
frm_def += 'Attribute VB_Creatable = False\n'
frm_def += 'Attribute VB_PredeclaredId = True\n'
frm_def += 'Attribute VB_Exposed = False\n'

# -- VB6 CODE -----------------------------------------------------------------
code = r"""Option Explicit
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
"""

content = frm_def + code

out = content.encode('windows-1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(out)
print('frmIBSCBS_Cadastro.frm gerado com sucesso')
