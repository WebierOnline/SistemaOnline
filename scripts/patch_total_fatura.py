# -*- coding: utf-8 -*-
# Corrige dois pontos relacionados ao valor liquido da fatura:
# 1. AtualizarTotaisNota: txtTotalFatura exibia varNota (com CBS/IBS) -> varProdutos - varDesconto
# 2. LerDadosInserir: ValorLiquidoFatura gravava txtTotaldaNota (com CBS/IBS) -> txtTotalFatura
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: AtualizarTotaisNota — txtTotalFatura ─────────────────────────────
patches.append(('txtTotalFatura',
    'txtTotalFatura.Text = FormatNumber(varNota, 2)\r\n',
    'txtTotalFatura.Text = FormatNumber(varProdutos - varDesconto, 2)\r\n',
    1))

# ── Patch 2: LerDadosInserir — ValorLiquidoFatura ────────────────────────────
patches.append(('ValorLiquidoFatura',
    '    TbNotas("ValorLiquidoFatura") = IIf(IsNull(Format(txtTotaldaNota, "@")) Or Vazio(Format(txtTotaldaNota, "@")), 0, CDbl(FormatNumber(txtTotaldaNota, 2)))\r\n',
    '    TbNotas("ValorLiquidoFatura") = IIf(IsNull(Format(txtTotalFatura, "@")) Or Vazio(Format(txtTotalFatura, "@")), 0, CDbl(FormatNumber(txtTotalFatura, 2)))\r\n',
    1))

all_ok = True
for name, old, new, expected in patches:
    c = text.count(old)
    ok = (c == expected)
    print(f'{name}: {"OK" if ok else "ERRO"} count={c} esperado={expected}')
    if not ok:
        all_ok = False

if all_ok:
    for _, old, new, _ in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('\nOK — txtTotalFatura e ValorLiquidoFatura corrigidos')
else:
    print('\nABORTADO')
