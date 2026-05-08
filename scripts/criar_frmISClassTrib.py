#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera C:\Projeto\OnlineCommerce\Forms\frmISClassTrib.frm
CRUD para tbISClassTrib — Imposto Seletivo: Classificações Tributárias

Estados dos botões:
  inicial    : Novo=ON,  Duplicar=OFF, Salvar=OFF,         Excluir=OFF
  novo       : Novo=OFF, Duplicar=OFF, Salvar=ON("Salvar"), Excluir=OFF
  selecionado: Novo=OFF, Duplicar=ON,  Salvar=ON("Editar"), Excluir=ON
  duplicar   : Novo=OFF, Duplicar=OFF, Salvar=ON("Salvar"), Excluir=OFF
"""

FORM_FILE = r"C:\Projeto\OnlineCommerce\Forms\frmISClassTrib.frm"

def hm(t):
    return int(round(t * 1.7639))

GW, GH = 13680, 3480

LAYOUT = f"""\
VERSION 5.00
Object = "{{5E9E78A0-531B-11CF-91F6-C2863C385E30}}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmISClassTrib
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IS - Classificações Tributárias"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblFiltroCod
      Caption         =   "Código IS:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label lblFiltroData
      Caption         =   "Vigente em:"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   180
      Width           =   1200
   End
   Begin VB.TextBox txtFiltroCod
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   150
      Width           =   1560
   End
   Begin VB.TextBox txtFiltroData
      Height          =   315
      Left            =   4260
      TabIndex        =   1
      Top             =   150
      Width           =   1560
   End
   Begin VB.CommandButton cmdFiltrar
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   5940
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid lstRegistros
      Height          =   {GH}
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   {GW}
      _ExtentX        =   {hm(GW)}
      _ExtentY        =   {hm(GH)}
      _Version        =   393216
   End
   Begin VB.Frame fraEdit
      Caption         =   "Dados"
      Height          =   2940
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   13680
      Begin VB.Label lblCod
         Caption         =   "Cód. IS:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   960
      End
      Begin VB.Label lblDescricao
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   300
         Width           =   960
      End
      Begin VB.Label lblTipo
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblISpAliq
         Caption         =   "Alíq. %:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label lblISvUnid
         Caption         =   "Val. Unid R$:"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   1140
         Width           =   1440
      End
      Begin VB.Label lblDIni
         Caption         =   "Vig. Ini:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label lblDFim
         Caption         =   "Vig. Fim:"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1560
         Width           =   960
      End
      Begin VB.TextBox txtCod
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox txtDescricao
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         Top             =   270
         Width           =   6120
      End
      Begin VB.ComboBox cboTipo
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   690
         Width           =   3600
      End
      Begin VB.TextBox txtISpAliq
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   1110
         Width           =   1440
      End
      Begin VB.TextBox txtISvUnid
         Height          =   315
         Left            =   4260
         TabIndex        =   9
         Top             =   1110
         Width           =   1440
      End
      Begin VB.TextBox txtDIni
         Height          =   315
         Left            =   1140
         TabIndex        =   12
         Top             =   1530
         Width           =   1440
      End
      Begin VB.TextBox txtDFim
         Height          =   315
         Left            =   3780
         TabIndex        =   13
         Top             =   1530
         Width           =   1440
      End
      Begin VB.CommandButton cmdNovo
         Caption         =   "Novo"
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   1440
      End
      Begin VB.CommandButton cmdDuplicar
         Caption         =   "Duplicar"
         Enabled         =   0  'False
         Height          =   450
         Left            =   1680
         TabIndex        =   15
         Top             =   2100
         Width           =   1440
      End
      Begin VB.CommandButton cmdSalvar
         Caption         =   "Salvar"
         Enabled         =   0  'False
         Height          =   450
         Left            =   3240
         TabIndex        =   16
         Top             =   2100
         Width           =   1440
      End
      Begin VB.CommandButton cmdExcluir
         Caption         =   "Excluir"
         Enabled         =   0  'False
         Height          =   450
         Left            =   4800
         TabIndex        =   17
         Top             =   2100
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmISClassTrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iId  As Long
Dim bNovo As Boolean

' ============================================================
' ESTADO DOS BOTÕES
' ============================================================
Private Sub DefinirEstado(sEstado As String)
    Select Case sEstado
        Case "inicial"
            cmdNovo.Enabled     = True
            cmdDuplicar.Enabled = False
            cmdSalvar.Enabled   = False
            cmdSalvar.Caption   = "Salvar"
            cmdExcluir.Enabled  = False
            LimparCampos
            iId = 0: bNovo = False

        Case "novo"
            cmdNovo.Enabled     = False
            cmdDuplicar.Enabled = False
            cmdSalvar.Enabled   = True
            cmdSalvar.Caption   = "Salvar"
            cmdExcluir.Enabled  = False
            LimparCampos
            iId = 0: bNovo = True
            cboTipo.ListIndex   = 0
            txtDIni.Text        = Format(Date, "dd/mm/yyyy")
            AtualizarCampos
            txtCod.SetFocus

        Case "selecionado"
            cmdNovo.Enabled     = False
            cmdDuplicar.Enabled = True
            cmdSalvar.Enabled   = True
            cmdSalvar.Caption   = "Editar"
            cmdExcluir.Enabled  = True
            bNovo = False

        Case "duplicar"
            cmdNovo.Enabled     = False
            cmdDuplicar.Enabled = False
            cmdSalvar.Enabled   = True
            cmdSalvar.Caption   = "Salvar"
            cmdExcluir.Enabled  = False
            iId = 0: bNovo = True
    End Select
End Sub

' ============================================================
' INICIALIZAÇÃO
' ============================================================
Private Sub Form_Load()
    With lstRegistros
        .Cols = 8
        .ColWidth(0) = 0
        .ColWidth(1) = 1020
        .ColWidth(2) = 4500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1080
        .ColWidth(5) = 1200
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .Row = 0: .Col = 0: .Text = "Id"
        .Row = 0: .Col = 1: .Text = "Cód. IS"
        .Row = 0: .Col = 2: .Text = "Descrição"
        .Row = 0: .Col = 3: .Text = "Tipo"
        .Row = 0: .Col = 4: .Text = "Alíq. %"
        .Row = 0: .Col = 5: .Text = "Val. Unid"
        .Row = 0: .Col = 6: .Text = "Vig. Ini"
        .Row = 0: .Col = 7: .Text = "Vig. Fim"
        .AllowUserResizing = 1
        .SelectionMode = 1
    End With
    cboTipo.Clear
    cboTipo.AddItem "1 - Ad Valorem (%)"
    cboTipo.AddItem "2 - Ad Rem (R$/unid)"
    cboTipo.AddItem "3 - Misto (% + R$/unid)"
    txtFiltroData.Text = Format(Date, "dd/mm/yyyy")
    DefinirEstado "inicial"
    CarregarRegistros
End Sub

' ============================================================
' GRID
' ============================================================
Private Sub CarregarRegistros()
    Dim rRs As ADODB.Recordset
    Dim sql As String, filtro As String, r As Long
    sql = "SELECT Id, cClassTrib_IS, Descricao, tipo_calculo_is, ISpAliq, ISvUnid, dIniVig, dFimVig FROM tbISClassTrib"
    If Trim(txtFiltroCod.Text) <> "" Then
        filtro = filtro & " AND cClassTrib_IS LIKE '%" & Trim(txtFiltroCod.Text) & "%'"
    End If
    If IsDate(txtFiltroData.Text) Then
        filtro = filtro & " AND CONVERT(date,'" & Format(CDate(txtFiltroData.Text), "yyyy-mm-dd") & "',23) BETWEEN dIniVig AND ISNULL(dFimVig,'9999-12-31')"
    End If
    If filtro <> "" Then sql = sql & " WHERE" & Mid(filtro, 5)
    sql = sql & " ORDER BY cClassTrib_IS, dIniVig"
    lstRegistros.rows = 1
    RsOpen rRs, sql
    Do While Not rRs.EOF
        lstRegistros.AddItem ""
        r = lstRegistros.rows - 1
        lstRegistros.Row = r: lstRegistros.Col = 0: lstRegistros.Text = CStr(rRs("Id"))
        lstRegistros.Row = r: lstRegistros.Col = 1: lstRegistros.Text = rRs("cClassTrib_IS")
        lstRegistros.Row = r: lstRegistros.Col = 2: lstRegistros.Text = rRs("Descricao")
        lstRegistros.Row = r: lstRegistros.Col = 3: lstRegistros.Text = TipoDesc(CInt(rRs("tipo_calculo_is")))
        lstRegistros.Row = r: lstRegistros.Col = 4: lstRegistros.Text = IIf(IsNull(rRs("ISpAliq")), "", Format(rRs("ISpAliq"), "0.00"))
        lstRegistros.Row = r: lstRegistros.Col = 5: lstRegistros.Text = IIf(IsNull(rRs("ISvUnid")), "", Format(rRs("ISvUnid"), "0.00"))
        lstRegistros.Row = r: lstRegistros.Col = 6: lstRegistros.Text = IIf(IsNull(rRs("dIniVig")), "", Format(rRs("dIniVig"), "dd/mm/yyyy"))
        lstRegistros.Row = r: lstRegistros.Col = 7: lstRegistros.Text = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "dd/mm/yyyy"))
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Function TipoDesc(t As Integer) As String
    Select Case t
        Case 1: TipoDesc = "Ad Valorem %"
        Case 2: TipoDesc = "Ad Rem R$/unid"
        Case 3: TipoDesc = "Misto"
        Case Else: TipoDesc = ""
    End Select
End Function

Private Sub lstRegistros_Click()
    If lstRegistros.Row = 0 Then Exit Sub
    iId = Val(lstRegistros.TextMatrix(lstRegistros.Row, 0))
    txtCod.Text      = lstRegistros.TextMatrix(lstRegistros.Row, 1)
    txtDescricao.Text = lstRegistros.TextMatrix(lstRegistros.Row, 2)
    Select Case lstRegistros.TextMatrix(lstRegistros.Row, 3)
        Case "Ad Valorem %":   cboTipo.ListIndex = 0
        Case "Ad Rem R$/unid": cboTipo.ListIndex = 1
        Case "Misto":          cboTipo.ListIndex = 2
        Case Else:             cboTipo.ListIndex = -1
    End Select
    txtISpAliq.Text = lstRegistros.TextMatrix(lstRegistros.Row, 4)
    txtISvUnid.Text = lstRegistros.TextMatrix(lstRegistros.Row, 5)
    txtDIni.Text    = lstRegistros.TextMatrix(lstRegistros.Row, 6)
    txtDFim.Text    = lstRegistros.TextMatrix(lstRegistros.Row, 7)
    DefinirEstado "selecionado"
    AtualizarCampos
End Sub

' ============================================================
' BOTÕES
' ============================================================
Private Sub cmdNovo_Click()
    DefinirEstado "novo"
End Sub

Private Sub cmdDuplicar_Click()
    DefinirEstado "duplicar"
End Sub

Private Sub cmdSalvar_Click()
    Dim sCod As String, sDesc As String, iTipo As Integer
    Dim sAliq As String, sUnid As String, sDIni As String, sDFim As String
    sCod  = Trim(txtCod.Text)
    sDesc = Trim(txtDescricao.Text)
    If cboTipo.ListIndex < 0 Then MsgBox "Selecione o tipo.", vbExclamation: Exit Sub
    iTipo = cboTipo.ListIndex + 1
    sAliq = Replace(Trim(txtISpAliq.Text), ",", ".")
    sUnid = Replace(Trim(txtISvUnid.Text), ",", ".")
    sDIni = Trim(txtDIni.Text)
    sDFim = Trim(txtDFim.Text)
    If sCod = "" Or sDesc = "" Or sDIni = "" Then
        MsgBox "Código, Descrição e Vig. Ini são obrigatórios.", vbExclamation: Exit Sub
    End If
    Dim sAliqSQL As String, sUnidSQL As String, sDFimSQL As String
    sAliqSQL = IIf(sAliq = "", "NULL", sAliq)
    sUnidSQL = IIf(sUnid = "", "NULL", sUnid)
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    Dim sql As String
    On Error GoTo ErrS
    If bNovo Then
        sql = "INSERT INTO tbISClassTrib " & _
              "(cClassTrib_IS, Descricao, tipo_calculo_is, ISpAliq, ISvUnid, dIniVig, dFimVig) " & _
              "VALUES ('" & sCod & "','" & Replace(sDesc, "'", "''") & "'," & _
              iTipo & "," & sAliqSQL & "," & sUnidSQL & "," & _
              "CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & ")"
    Else
        sql = "UPDATE tbISClassTrib SET " & _
              "cClassTrib_IS='" & sCod & "'," & _
              "Descricao='" & Replace(sDesc, "'", "''") & "'," & _
              "tipo_calculo_is=" & iTipo & "," & _
              "ISpAliq=" & sAliqSQL & "," & _
              "ISvUnid=" & sUnidSQL & "," & _
              "dIniVig=CONVERT(date,'" & sDIni & "',103)," & _
              "dFimVig=" & sDFimSQL & " WHERE Id=" & iId
    End If
    vgDb.Execute sql
    MsgBox "Salvo.", vbInformation
    CarregarRegistros
    DefinirEstado "inicial"
    Exit Sub
ErrS: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluir_Click()
    If iId = 0 Then MsgBox "Selecione um registro.", vbExclamation: Exit Sub
    If MsgBox("Excluir este registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrD
    vgDb.Execute "DELETE FROM tbISClassTrib WHERE Id=" & iId
    MsgBox "Excluído.", vbInformation
    CarregarRegistros
    DefinirEstado "inicial"
    Exit Sub
ErrD: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ============================================================
' AUXILIARES
' ============================================================
Private Sub cboTipo_Click()
    AtualizarCampos
End Sub

Private Sub AtualizarCampos()
    If cboTipo.ListIndex < 0 Then Exit Sub
    Dim t As Integer
    t = cboTipo.ListIndex + 1
    txtISpAliq.Enabled = (t = 1 Or t = 3)
    txtISvUnid.Enabled = (t = 2 Or t = 3)
    If Not txtISpAliq.Enabled Then txtISpAliq.Text = ""
    If Not txtISvUnid.Enabled Then txtISvUnid.Text = ""
End Sub

Private Sub LimparCampos()
    txtCod.Text       = ""
    txtDescricao.Text = ""
    cboTipo.ListIndex = -1
    txtISpAliq.Text   = ""
    txtISvUnid.Text   = ""
    txtDIni.Text      = ""
    txtDFim.Text      = ""
End Sub

Private Sub cmdFiltrar_Click()
    CarregarRegistros
    DefinirEstado "inicial"
End Sub

Private Sub txtFiltroCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CarregarRegistros
End Sub

Private Sub txtFiltroData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CarregarRegistros
End Sub
"""

content = LAYOUT.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
with open(FORM_FILE, 'wb') as f:
    f.write(content.encode('cp1252'))

print(f"Gravado: {FORM_FILE}")
