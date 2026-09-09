# -*- coding: utf-8 -*-
"""
Feature: categoria padrao.
- Categorias_Cadastro.frm: botao cmdPadrao define/limpa a categoria padrao do Tipo_Empresa,
  grid ganha coluna PADRAO.
- Produtos_Cadastro.frm: produto novo ja vem com a categoria padrao selecionada;
  cboCategoria_GotFocus preserva a selecao ao repopular.
Ambos .frm sao cp1252 -> edicao binaria + normalizacao de CRLF.
"""
import io, sys

def norm(b: bytes) -> bytes:
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

def patch(path, repls):
    with open(path, "rb") as f:
        data = f.read()
    data = data.replace(b"\r\n", b"\n")  # trabalha em LF, normaliza no fim
    for i, (old, new) in enumerate(repls, 1):
        if new in data:
            print(f"  [{path}] repl #{i}: ja aplicado, pulando")
            continue
        n = data.count(old)
        if n != 1:
            print(f"  [{path}] repl #{i}: esperado 1 ocorrencia, achou {n}")
            print("  --- OLD ---")
            print(old.decode("cp1252", "replace"))
            sys.exit(1)
        data = data.replace(old, new)
    data = norm(data)
    with open(path, "wb") as f:
        f.write(data)
    print(f"  OK {path} ({len(repls)} trecho(s))")


# ---------------------------------------------------------------- Categorias_Cadastro.frm
cat = r"C:\projeto\Compartilhado\Forms\Categorias_Cadastro.frm"

cat_repls = [
    # 1) habilita/desabilita cmdPadrao junto com os demais botoes de linha
    (
b'''Private Sub HabilitarEdicao()
    txtCategoria.Enabled = True
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNovo.Enabled = False
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    lblAviso.Caption = ""
End Sub

Private Sub DesabilitarEdicao()
    txtCategoria.Text = ""
    txtCategoria.Enabled = False
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNovo.Enabled = True
    cmdEditar.Enabled = True
    cmdExcluir.Enabled = True
    vIDCategoria = 0
    vTipoEdicao = ""
    lblAviso.Caption = ""
End Sub''',
b'''Private Sub HabilitarEdicao()
    txtCategoria.Enabled = True
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNovo.Enabled = False
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    cmdPadrao.Enabled = False
    lblAviso.Caption = ""
End Sub

Private Sub DesabilitarEdicao()
    txtCategoria.Text = ""
    txtCategoria.Enabled = False
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNovo.Enabled = True
    cmdEditar.Enabled = True
    cmdExcluir.Enabled = True
    cmdPadrao.Enabled = True
    vIDCategoria = 0
    vTipoEdicao = ""
    lblAviso.Caption = ""
End Sub''',
    ),
    # 2) query do grid traz Padrao
    (
b'    RsOpen r, "SELECT ID_Categoria, Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria"',
b'    RsOpen r, "SELECT ID_Categoria, Categoria, Padrao FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria"',
    ),
    # 3) FormatarGrid: 3 colunas
    (
b'        .Cols = 2\n        .rows = 2',
b'        .Cols = 3\n        .rows = 2',
    ),
    (
b'        .ColWidth(0) = 600\n        .ColWidth(1) = 7740\n',
b'        .ColWidth(0) = 600\n        .ColWidth(1) = 6540\n        .ColWidth(2) = 1200\n',
    ),
    (
b'''        .TextMatrix(0, 1) = "CATEGORIA"
        .Col = 0: .Row = 0: .CellFontBold = True: .CellAlignment = 4
        .Col = 1: .Row = 0: .CellFontBold = True: .CellAlignment = 4
''',
b'''        .TextMatrix(0, 1) = "CATEGORIA"
        .TextMatrix(0, 2) = "PADR\xc3O"
        .Col = 0: .Row = 0: .CellFontBold = True: .CellAlignment = 4
        .Col = 1: .Row = 0: .CellFontBold = True: .CellAlignment = 4
        .Col = 2: .Row = 0: .CellFontBold = True: .CellAlignment = 4
''',
    ),
    (
b'''                .TextMatrix(.rows - 1, 0) = rTabela("ID_Categoria")
                .TextMatrix(.rows - 1, 1) = rTabela("Categoria")
                rTabela.MoveNext''',
b'''                .TextMatrix(.rows - 1, 0) = rTabela("ID_Categoria")
                .TextMatrix(.rows - 1, 1) = rTabela("Categoria")
                If Not IsNull(rTabela("Padrao")) Then
                    If rTabela("Padrao") Then
                        .TextMatrix(.rows - 1, 2) = "SIM"
                        .Col = 2: .Row = .rows - 1: .CellFontBold = True: .CellAlignment = 4
                    End If
                End If
                rTabela.MoveNext''',
    ),
    # 4) handler do cmdPadrao (toggle: define a selecionada e zera as demais; se ja era, so limpa)
    (
b'''Private Sub cmdCancelar_Click()
    DesabilitarEdicao
End Sub''',
b'''Private Sub cmdPadrao_Click()
    If gridCategorias.Row < 1 Then
        MsgBox "Selecione uma categoria para definir como padr\xe3o!", vbExclamation, "Aten\xe7\xe3o"
        Exit Sub
    End If
    Dim vID As Long, jaEraPadrao As Boolean
    vID = CLng(gridCategorias.TextMatrix(gridCategorias.Row, 0))
    jaEraPadrao = (Trim(gridCategorias.TextMatrix(gridCategorias.Row, 2)) <> "")
    SQLExecuta "UPDATE Categorias SET Padrao = 0 WHERE Tipo_Empresa = " & tipoEmpresa
    If Not jaEraPadrao Then
        SQLExecuta "UPDATE Categorias SET Padrao = 1 WHERE ID_Categoria = " & vID
    End If
    ExibirGrid
End Sub

Private Sub cmdCancelar_Click()
    DesabilitarEdicao
End Sub''',
    ),
]
patch(cat, cat_repls)


# ---------------------------------------------------------------- Produtos_Cadastro.frm
prd = r"C:\projeto\Compartilhado\Forms\Produtos_Cadastro.frm"

prd_repls = [
    # 1) GotFocus preserva selecao + nova sub PreencherCategoriaPadrao
    (
b'''Private Sub cboCategoria_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
cboCategoria.Clear
sSQL = "SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria"
Set r = dbData.OpenRecordset(sSQL)
Do While Not r.EOF
   cboCategoria.AddItem ValidateNull(r("Categoria"))
   r.MoveNext
Loop
If r.State <> 0 Then r.Close
End Sub''',
b'''Private Sub cboCategoria_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim sSelAtual As String
sSelAtual = cboCategoria.Text
cboCategoria.Clear
sSQL = "SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria"
Set r = dbData.OpenRecordset(sSQL)
Do While Not r.EOF
   cboCategoria.AddItem ValidateNull(r("Categoria"))
   r.MoveNext
Loop
If r.State <> 0 Then r.Close
If sSelAtual <> "" Then SelecionarNoCombo cboCategoria, sSelAtual
End Sub

Private Sub PreencherCategoriaPadrao()
' Produto novo: popula o combo e ja seleciona a categoria marcada como padrao (se houver).
' Sem categoria padrao cadastrada -> combo fica vazio esperando a escolha.
Dim rC As ADODB.Recordset
Dim rP As ADODB.Recordset
Dim sPad As String
cboCategoria.Clear
Set rC = dbData.OpenRecordset("SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria")
Do While Not rC.EOF
   cboCategoria.AddItem ValidateNull(rC("Categoria"))
   rC.MoveNext
Loop
If rC.State <> 0 Then rC.Close
On Error Resume Next   ' base ainda sem a coluna Padrao (script 111 nao rodado) -> segue sem padrao
Set rP = dbData.OpenRecordset("SELECT TOP 1 Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " AND Padrao = 1")
If Err.Number = 0 Then
   If Not rP.EOF Then sPad = ValidateNull(rP("Categoria"))
   If rP.State <> 0 Then rP.Close
End If
On Error GoTo 0
If sPad <> "" Then SelecionarNoCombo cboCategoria, sPad
End Sub''',
    ),
    # 2) chama no cmdNovo_Click
    (
b'''cboUnidMedida.Text = "UN"
txtQuant.Text = "0"
SelecionarNoCombo cboISCST, "99", True
txtCodBarra.SetFocus
End Sub''',
b'''cboUnidMedida.Text = "UN"
txtQuant.Text = "0"
SelecionarNoCombo cboISCST, "99", True
PreencherCategoriaPadrao
txtCodBarra.SetFocus
End Sub''',
    ),
]
patch(prd, prd_repls)

print("Feito.")
