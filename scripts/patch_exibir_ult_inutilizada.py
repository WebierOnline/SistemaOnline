# -*- coding: utf-8 -*-
# Corrige ExibirUltimasNfe: adiciona caso Inutilizada no SQL e loop de contagem
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: SQL — adicionar Inutilizada no CASE ─────────────────────────────
patches.append(('sql_inutilizada',
    '"(CASE WHEN Denegada = 1 THEN \'Denegada\' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN \'Enviada\' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN \'Cancelada\' ELSE \'Em Digita\xe7\xe3o\' END) END) END) AS Status " & _\r\n'
    '                    "FROM NotaFiscal order by NumeroNota desc"',

    '"(CASE WHEN Denegada = 1 THEN \'Denegada\' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 AND Inutilizada = 0 THEN \'Enviada\' ELSE (CASE WHEN Enviada = 1 AND Inutilizada = 1 THEN \'Inutilizada\' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN \'Cancelada\' ELSE \'Em Digita\xe7\xe3o\' END) END) END) END) AS Status " & _\r\n'
    '                    "FROM NotaFiscal order by NumeroNota desc"',
))

# ── Patch 2: adicionar loop Inutilizada e usar variaveis no ListView ──────────
patches.append(('loop_inutilizada',
    'Dim oLI As Object\r\n'
    'With lvwTotais\r\n'
    '    .ListItems.Clear\r\n'
    '    Set oLI = .ListItems.Add(, , "Enviadas")\r\n'
    '    oLI.SubItems(1) = Format(nEnv, "000")\r\n'
    '    oLI.SubItems(2) = Format(tEnv, ocMONEY)\r\n'
    '    Set oLI = .ListItems.Add(, , "N\xe3o Enviadas")\r\n'
    '    oLI.SubItems(1) = Format(contar, "000")\r\n'
    '    oLI.SubItems(2) = Format(soma, ocMONEY)\r\n'
    '    Set oLI = .ListItems.Add(, , "Canceladas")\r\n'
    '    oLI.SubItems(1) = Format(nCan, "000")\r\n'
    '    oLI.SubItems(2) = Format(tCan, ocMONEY)\r\n'
    '    Set oLI = .ListItems.Add(, , "Inutilizadas")\r\n'
    '    oLI.SubItems(1) = "000"\r\n'
    '    oLI.SubItems(2) = Format(0, ocMONEY)\r\n'
    'End With\r\n'
    '\r\n'
    '\r\n'
    'Exit Sub\r\n'
    'Resume\r\n'
    'End Sub',

    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Inutilizada" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'Dim nInu As Long, tInu As Currency\r\n'
    'nInu = contar: tInu = soma\r\n'
    '\r\n'
    'Dim oLI As Object\r\n'
    'With lvwTotais\r\n'
    '    .ListItems.Clear\r\n'
    '    Set oLI = .ListItems.Add(, , "Enviadas")\r\n'
    '    oLI.SubItems(1) = Format(nEnv, "000")\r\n'
    '    oLI.SubItems(2) = Format(tEnv, ocMONEY)\r\n'
    '    Set oLI = .ListItems.Add(, , "N\xe3o Enviadas")\r\n'
    '    oLI.SubItems(1) = Format(contar, "000")\r\n'
    '    oLI.SubItems(2) = Format(soma, ocMONEY)\r\n'
    '    Set oLI = .ListItems.Add(, , "Canceladas")\r\n'
    '    oLI.SubItems(1) = Format(nCan, "000")\r\n'
    '    oLI.SubItems(2) = Format(tCan, ocMONEY)\r\n'
    '    Set oLI = .ListItems.Add(, , "Inutilizadas")\r\n'
    '    oLI.SubItems(1) = Format(nInu, "000")\r\n'
    '    oLI.SubItems(2) = Format(tInu, ocMONEY)\r\n'
    'End With\r\n'
    '\r\n'
    '\r\n'
    'Exit Sub\r\n'
    'Resume\r\n'
    'End Sub',
))

all_ok = True
for name, old, new in patches:
    c = text.count(old)
    status = 'OK' if c == 1 else f'ERRO count={c}'
    print(f'{name}: {status}')
    if c != 1:
        all_ok = False

if all_ok:
    for _, old, new in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('\nOK')
else:
    print('\nABORTADO')
