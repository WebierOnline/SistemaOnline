# -*- coding: utf-8 -*-
"""
frmBuscarPlaca.frm: adiciona cod_os como coluna oculta (9) no grid, para
que cada linha represente uma OS especifica (nao mais um veiculo
"agrupado"). cmdUsarEsse_Click passa a capturar tambem lCodOSSelecionado,
usado pelo novo Menu_Consulta_Placa em OS_Recapadora para carregar a OS
exata clicada.
"""

PATH = r"C:\projeto\OrdemServico\Forms\frmBuscarPlaca.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


# 1) Public lCodOSSelecionado
i = find_line_exact("Public sChassiSel As String")
lines.insert(i + 1, "Public lCodOSSelecionado As Long")

# 2) ConfigurarGrid: Cols 9 -> 10, ColWidth(9) = 0
i = find_line_exact("        .Cols = 9")
lines[i] = "        .Cols = 10"
i = find_line_exact('        .Row = 0: .Col = 2: .Text = "CLIENTE"')
lines.insert(i, "        .ColWidth(9) = 0")

# 3) CarregarGrid: SQL + coluna 9
old_sql = [
    '    sql = "SELECT DISTINCT cliente.codigo AS cod_cliente, cliente.nome, cliente.celular, " & _',
    '          "OS_Equipamento_Auto.modelo, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.placa, " & _',
    '          "OS_Equipamento_Auto.km, OS_Equipamento_Auto.cor, OS_Equipamento_Auto.chassi " & _',
    '          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE " & _',
    '          "INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _',
    '          "WHERE (OS_Equipamento_Auto.placa LIKE \'%" & Replace(Trim(txtPlacaF.Text), "\'", "\'\'") & "%\') " & _',
    '          "ORDER BY cliente.nome"',
]
i_sql = find_line_exact(old_sql[0])
for k, l in enumerate(old_sql):
    assert lines[i_sql + k] == l, (i_sql + k, repr(lines[i_sql + k]), repr(l))

new_sql = [
    '    sql = "SELECT DISTINCT cliente.codigo AS cod_cliente, cliente.nome, cliente.celular, " & _',
    '          "OS_Equipamento_Auto.modelo, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.placa, " & _',
    '          "OS_Equipamento_Auto.km, OS_Equipamento_Auto.cor, OS_Equipamento_Auto.chassi, OS.COD_OS AS cod_os " & _',
    '          "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE " & _',
    '          "INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _',
    '          "WHERE (OS_Equipamento_Auto.placa LIKE \'%" & Replace(Trim(txtPlacaF.Text), "\'", "\'\'") & "%\') " & _',
    '          "ORDER BY cliente.nome"',
]
lines[i_sql : i_sql + len(old_sql)] = new_sql

i = find_line_exact('            .Row = n: .Col = 8: .Text = ValidateNull(rVei("chassi"))')
lines.insert(i + 1, '            .Row = n: .Col = 9: .Text = ValidateNull(rVei("cod_os"))')

# 4) cmdUsarEsse_Click: captura lCodOSSelecionado
i = find_line_exact('    sChassiSel = lstVeiculos.TextMatrix(lstVeiculos.Row, 8)')
lines.insert(i + 1, '    lCodOSSelecionado = Val(lstVeiculos.TextMatrix(lstVeiculos.Row, 9))')

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - cod_os adicionado ao grid e captura de lCodOSSelecionado")
