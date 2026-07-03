# -*- coding: utf-8 -*-
"""
Patch: Vendas_Consulta_PorProdutos.frm — POR SERVICOS passa a usar OS.DATA_TERMINO
no lugar de OS_Servicos_Auto.data (alias s.data) para filtros e exibicao.

Mudancas:
  1. sBase: s.data AS varData  ->  OS.DATA_TERMINO AS varData
  2. WHEREs MENSAL:     MONTH(s.data) / YEAR(s.data)
  3. WHEREs PERIODO:    (s.data >= ...) / (s.data <= ...)
  4. WHEREs SERVICOS/MENSAL e SERVICOS/PERIODO: idem
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# 1. sBase alias — unico no arquivo
changes.append(('s.data AS varData',
                'OS.DATA_TERMINO AS varData',
                'sBase alias varData', 1))

# 2. MONTH(s.data) — aparece em MENSAL(3x) + SERVICOS/MENSAL(2x) = 5x
changes.append(('MONTH(s.data)',
                'MONTH(OS.DATA_TERMINO)',
                'MONTH(s.data) -> MONTH(OS.DATA_TERMINO)', 5))

# 3. YEAR(s.data) — mesmos 5 blocos
changes.append(('YEAR(s.data)',
                'YEAR(OS.DATA_TERMINO)',
                'YEAR(s.data) -> YEAR(OS.DATA_TERMINO)', 5))

# 4. (s.data >= CONVERT(DATETIME — PERIODO(3x) + SERVICOS/PERIODO(2x) = 5x
changes.append(('(s.data >= CONVERT(DATETIME',
                '(OS.DATA_TERMINO >= CONVERT(DATETIME',
                '(s.data >= ...) -> (OS.DATA_TERMINO >= ...)', 5))

# 5. (s.data <= CONVERT(DATETIME — mesmos 5 blocos
changes.append(('(s.data <= CONVERT(DATETIME',
                '(OS.DATA_TERMINO <= CONVERT(DATETIME',
                '(s.data <= ...) -> (OS.DATA_TERMINO <= ...)', 5))

for old, new, label, expected in changes:
    count = text.count(old)
    if count != expected:
        print(f'ERRO [{label}]: {count} ocorrencias (esperado {expected})')
        sys.exit(1)
    text = text.replace(old, new)
    print(f'OK ({count}x): {label}')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
