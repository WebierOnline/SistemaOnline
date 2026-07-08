"""
Patch: Vendas_Consulta_PorProdutos.frm — bloco POR PRODUTOS
1. SELECT: pedidos_itens.data as varData -> pedidos.DATA_COMPRA as varData
2. Todos os filtros MENSAL/PERIODO/DATA: pedidos_itens.data -> pedidos.DATA_COMPRA (31 ocorrencias)
3. Adiciona casos faltantes: TODOS+MENSAL, TODOS+PERIODO, TODOS+DATA

Motivacao: sem os casos TODOS+*, nenhum filtro de data era aplicado
e a query retornava todos os registros de todos os tempos.
"""

FRM = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FRM, 'rb') as f:
    raw = f.read()

data = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = data.decode('windows-1252')

ok = True

# --- Patch 1: replace_all pedidos_itens.data -> pedidos.DATA_COMPRA ---
count1 = text.count('pedidos_itens.data')
if count1 != 31:
    print(f'ERRO [replace_all]: {count1} ocorrencias (esperado 31)')
    ok = False
else:
    print(f'OK [replace_all pedidos_itens.data]: {count1} ocorrencias')
    text = text.replace('pedidos_itens.data', 'pedidos.DATA_COMPRA')

if not ok:
    print('Abortando.')
    exit()

# --- Patch 2: Inserir caso TODOS+MENSAL antes do comentario 'PERIODO ---
old2 = (
    '= " & cboAno & ") " & _\n'
    '                       "ORDER BY " & INDICE\n'
    "            'PERÍODO\n"
    '             ElseIf cboCriterioSec.Text ='
)
new2 = (
    '= " & cboAno & ") " & _\n'
    '                       "ORDER BY " & INDICE\n'
    "            'TODOS/MENSAL\n"
    '             ElseIf cboCriterioSec.Text = "TODOS" And cboCriterioPrinc.Text = "MENSAL" Then\n'
    '                If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub\n'
    '                sSQL = sSQL & " and (MONTH(pedidos.DATA_COMPRA) = " & cboMes.ListIndex + 1 & ") AND (YEAR(pedidos.DATA_COMPRA) = " & cboAno & ") " & _\n'
    '                       "ORDER BY " & INDICE\n'
    "            'PERÍODO\n"
    '             ElseIf cboCriterioSec.Text ='
)
c2 = text.count(old2)
if c2 != 1:
    print(f'ERRO [TODOS/MENSAL]: {c2} ocorrencias (esperado 1)')
    ok = False
else:
    print('OK [TODOS/MENSAL]: encontrado')
    text = text.replace(old2, new2)

# --- Patch 3: Inserir caso TODOS+PERIODO antes do comentario 'DATA ---
old3 = (
    "', 103)) \" & _\n"
    '                       "ORDER BY " & INDICE\n'
    '\n'
    "            'DATA\n"
    '             ElseIf cboCriterioSec.Text = "D'
)
new3 = (
    "', 103)) \" & _\n"
    '                       "ORDER BY " & INDICE\n'
    "            'TODOS/PERÍODO\n"
    '             ElseIf cboCriterioSec.Text = "TODOS" And cboCriterioPrinc.Text = "PERÍODO" Then\n'
    '                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub\n'
    '                sSQL = sSQL & " and (pedidos.DATA_COMPRA >= CONVERT(DATETIME, \'" & Format(mskInicio.Text, ocDATA) & "\', 103)) AND (pedidos.DATA_COMPRA <= CONVERT(DATETIME, \'" & Format(mskFim.Text, ocDATA) & "\', 103)) " & _\n'
    '                       "ORDER BY " & INDICE\n'
    '\n'
    "            'DATA\n"
    '             ElseIf cboCriterioSec.Text = "D'
)
c3 = text.count(old3)
if c3 != 1:
    print(f'ERRO [TODOS/PERIODO]: {c3} ocorrencias (esperado 1)')
    ok = False
else:
    print('OK [TODOS/PERIODO]: encontrado')
    text = text.replace(old3, new3)

# --- Patch 4: Inserir caso TODOS+DATA antes do comentario 'PRODUTO/MENSAL ---
old4 = (
    "', 103)) \" & _\n"
    '                       "ORDER BY " & INDICE\n'
    "            'PRODUTO/MENSAL\n"
    '             ElseIf cboCriterioPri'
)
new4 = (
    "', 103)) \" & _\n"
    '                       "ORDER BY " & INDICE\n'
    "            'TODOS/DATA\n"
    '             ElseIf cboCriterioSec.Text = "TODOS" And cboCriterioPrinc.Text = "DATA" Then\n'
    '                If Not IsDate(mskInicio.Text) Then Exit Sub\n'
    '                sSQL = sSQL & " and (pedidos.DATA_COMPRA = CONVERT(DATETIME, \'" & Format(mskInicio.Text, ocDATA) & "\', 103)) " & _\n'
    '                       "ORDER BY " & INDICE\n'
    "            'PRODUTO/MENSAL\n"
    '             ElseIf cboCriterioPri'
)
c4 = text.count(old4)
if c4 != 1:
    print(f'ERRO [TODOS/DATA]: {c4} ocorrencias (esperado 1)')
    ok = False
else:
    print('OK [TODOS/DATA]: encontrado')
    text = text.replace(old4, new4)

if not ok:
    print('Erros encontrados — arquivo NAO salvo.')
else:
    result = text.encode('windows-1252').replace(b'\n', b'\r\n')
    with open(FRM, 'wb') as f:
        f.write(result)
    print('Arquivo salvo com sucesso.')
