# -*- coding: utf-8 -*-
"""Etiquetas_Impressao.frm: filtro de descricao por Palavra / Palavras Duplas
(mesma logica de Produtos_Cadastro). Remove chkDescPorProduto / chkDescPorIniciais.
.frm cp1252 -> edicao binaria + CRLF."""
import sys

p = r"C:\projeto\Compartilhado\Forms\Etiquetas_Impressao.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n")

reps = []

# --- .frm: optPorPalavra/optPalavrasDuplas comecam invisiveis (so no modo Descricao) ---
reps.append((
b'''      Begin VB.OptionButton optPalavrasDuplas 
         Caption         =   "Palavras Duplas"
         Height          =   195
         Left            =   1200
         TabIndex        =   45
         Top             =   1320
         Width           =   1515
      End
      Begin VB.OptionButton optPorPalavra 
         Caption         =   "Palavra"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   915
      End
''',
b'''      Begin VB.OptionButton optPalavrasDuplas 
         Caption         =   "Palavras Duplas"
         Height          =   195
         Left            =   1200
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optPorPalavra 
         Caption         =   "Palavra"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   915
      End
'''))

# --- .frm: remove os 2 checkboxes ---
reps.append((
b'''      Begin VB.CheckBox chkDescPorIniciais 
         Caption         =   "Por Iniciais"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1500
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkDescPorProduto 
         Caption         =   "Por Produto"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1275
      End
''',
b''))

# --- codigo: remove as 2 Subs de evento dos checkboxes ---
reps.append((
b'''Private Sub chkDescPorIniciais_Click()
   If optDesc.Value = Unchecked Then Exit Sub

   If chkDescPorIniciais.Value = Checked Then
      cboDesc.Clear
      chkDescPorProduto.Value = Unchecked
      cboDesc.SetFocus
   End If
End Sub

Private Sub chkDescPorProduto_Click()
   If optDesc.Value = Unchecked Then Exit Sub

   If chkDescPorProduto.Value = Checked Then
      chkDescPorIniciais.Value = Unchecked
      cboDesc.SetFocus
   End If
End Sub

''',
b''))

# --- MostrarCriterios: filtro de descricao por Palavra / Palavras Duplas ---
reps.append((
b'''   If chkDescPorProduto.Value = Checked Then
      var_Criterio = var_Criterio & IIf(optDesc.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao = '" & cboDesc.Text & "'", "")
   ElseIf chkDescPorIniciais.Value = Checked Then
      var_Criterio = Chr$(39) & cboDesc.Text & "%" & Chr(39)
      var_Criterio = var_Criterio & IIf(optDesc.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao  LIKE " & var_Criterio & "", "")
   End If
''',
b'''   If optDesc.Value = True Then
      Dim vDescFiltro As String
      Dim aPalDesc() As String
      Dim iPalDesc As Integer
      vDescFiltro = ""
      If optPorPalavra.Value = True Then
         vDescFiltro = "produtos.descricao LIKE '%" & cboDesc.Text & "%'"
      ElseIf optPalavrasDuplas.Value = True Then
         aPalDesc = Split(Trim(cboDesc.Text), " ")
         For iPalDesc = 0 To UBound(aPalDesc)
            If Trim(aPalDesc(iPalDesc)) <> "" Then
               If vDescFiltro <> "" Then vDescFiltro = vDescFiltro & " AND "
               vDescFiltro = vDescFiltro & "produtos.descricao LIKE '%" & Trim(aPalDesc(iPalDesc)) & "%'"
            End If
         Next iPalDesc
      End If
      If vDescFiltro <> "" Then var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(" & vDescFiltro & ")"
   End If
'''))

# --- cboDesc_GotFocus: guard antigo era chkDescPorProduto ---
reps.append((
b'''   If chkDescPorProduto.Value = Checked Then
      cboDesc.Clear
''',
b'''   If optDesc.Value = True Then
      cboDesc.Clear
'''))

# --- optDesc_Click: mostra optPorPalavra/optPalavrasDuplas (default Palavra) ---
reps.append((
b'''   lblDesc.Visible = True
   cboDesc.Visible = True
   chkDescPorProduto.Visible = True
   chkDescPorIniciais.Visible = True
   lblCodBarra.Visible = False
''',
b'''   lblDesc.Visible = True
   cboDesc.Visible = True
   optPorPalavra.Visible = True
   optPalavrasDuplas.Visible = True
   optPorPalavra.Value = True
   lblCodBarra.Visible = False
'''))

# --- optCategoria_Click / optCodBarra_Click / optTodos_Click: esconde os 2 opts (3 ocorrencias) ---
old_hide = b'   chkDescPorProduto.Visible = False\n   chkDescPorIniciais.Visible = False\n'
new_hide = b'   optPorPalavra.Visible = False\n   optPalavrasDuplas.Visible = False\n'

for i, (o, n) in enumerate(reps, 1):
    if (n == b'' and o not in d) or (n != b'' and n in d and o not in d):
        print(f"#{i} ja aplicado"); continue
    if d.count(o) != 1:
        print(f"#{i}: {d.count(o)} ocorrencias (esperado 1) -- ABORTA"); sys.exit(1)
    d = d.replace(o, n, 1); print(f"#{i} OK")

c = d.count(old_hide)
if c == 3:
    d = d.replace(old_hide, new_hide); print("#hide OK (3x)")
elif new_hide in d and old_hide not in d:
    print("#hide ja aplicado")
else:
    print(f"#hide: {c} ocorrencias (esperado 3) -- ABORTA"); sys.exit(1)

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado")
