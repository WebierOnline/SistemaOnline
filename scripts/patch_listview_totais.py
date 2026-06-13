# -*- coding: utf-8 -*-
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

# ── Patch A: Substituir Frame2 (12 labels) por MSComctlLib.ListView ──────────
lines = text.split('\r\n')

assert 'Begin VB.Frame Frame2 ' in lines[435], f'Linha 435 inesperada: {repr(lines[435])}'
assert lines[680] == '      End', f'Linha 680 inesperada: {repr(lines[680])}'

listview_lines = [
    '      Begin MSComctlLib.ListView lvwTotais ',
    '         Height          =   1155',
    '         Left            =   13080',
    '         TabIndex        =   263',
    '         Top             =   6360',
    '         Width           =   3435',
    '         _ExtentX        =   6059',
    '         _ExtentY        =   2037',
    '         View            =   3',
    '         LabelEdit       =   1',
    "         LabelWrap       =   -1  'True",
    "         HideSelection   =   -1  'True",
    "         FullRowSelect   =   -1  'True",
    "         GridLines       =   -1  'True",
    '         _Version        =   393217',
    '         ForeColor       =   -2147483640',
    '         BackColor       =   -2147483643',
    '         BorderStyle     =   1',
    '         Appearance      =   1',
    '         NumItems        =   0',
    '      End',
]

lines = lines[:435] + listview_lines + lines[681:]
text = '\r\n'.join(lines)

patches = []

# ── Patch B: Inicializar colunas do ListView no Form_Load ────────────────────
patches.append(('form_load_colunas',
    'Set moCombo = New cComboHelper\r\n'
    'Frm_NF.Tab = 0',

    'Set moCombo = New cComboHelper\r\n'
    'With lvwTotais.ColumnHeaders\r\n'
    '    .Add , , "Status", 1320\r\n'
    '    .Add , , "Qtd", 510\r\n'
    '    .Add , , "Valor", 1440\r\n'
    'End With\r\n'
    'Frm_NF.Tab = 0',
))

# ── Patch C: ExibirUltimasNfe — substituir 6 lblXxx por ListView ─────────────
# Ancora unica: bloco inteiro ate Exit Sub/Resume/End Sub
patches.append(('exibir_ult_nfe',
    'lblQuantEnviada.Caption = Format(contar, "000")\r\n'
    'lblTotalEnviada.Caption = Format(soma, ocMONEY)\r\n'
    '\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Cancelada" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'lblQuantCancelada.Caption = Format(contar, "000")\r\n'
    'lblTotalCancelada.Caption = Format(soma, ocMONEY)\r\n'
    '\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Em Digita\xe7\xe3o" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'lblQuantNaoEnviada.Caption = Format(contar, "000")\r\n'
    'lblTotalNaoEnviada.Caption = Format(soma, ocMONEY)\r\n'
    '\r\n'
    '\r\n'
    'Exit Sub\r\n'
    'Resume\r\n'
    'End Sub',

    'Dim nEnv As Long, tEnv As Currency\r\n'
    'nEnv = contar: tEnv = soma\r\n'
    '\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Cancelada" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'Dim nCan As Long, tCan As Currency\r\n'
    'nCan = contar: tCan = soma\r\n'
    '\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Em Digita\xe7\xe3o" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
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
    '    oLI.SubItems(1) = "000"\r\n'
    '    oLI.SubItems(2) = Format(0, ocMONEY)\r\n'
    'End With\r\n'
    '\r\n'
    '\r\n'
    'Exit Sub\r\n'
    'Resume\r\n'
    'End Sub',
))

# ── Patch D: cmdExibirConNotas_Click — substituir 8 lblXxx por ListView ───────
# Ancora unica: contem lblQuantInutilizada (count=1 no arquivo)
patches.append(('exibir_con_notas',
    'lblQuantEnviada.Caption = Format(contar, "000")\r\n'
    'lblTotalEnviada.Caption = Format(soma, ocMONEY)\r\n'
    '\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Cancelada" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'lblQuantCancelada.Caption = Format(contar, "000")\r\n'
    'lblTotalCancelada.Caption = Format(soma, ocMONEY)\r\n'
    '\r\n'
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
    'lblQuantInutilizada.Caption = Format(contar, "000")\r\n'
    'lblTotalInutilizada.Caption = Format(soma, ocMONEY)\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Em Digita\xe7\xe3o" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'lblQuantNaoEnviada.Caption = Format(contar, "000")\r\n'
    'lblTotalNaoEnviada.Caption = Format(soma, ocMONEY)\r\n'
    '\r\n'
    'cmdEditar.Caption = "Editar"',

    'Dim nEnv As Long, tEnv As Currency\r\n'
    'nEnv = contar: tEnv = soma\r\n'
    '\r\n'
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Cancelada" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
    '\r\n'
    'Dim nCan As Long, tCan As Currency\r\n'
    'nCan = contar: tCan = soma\r\n'
    '\r\n'
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
    "'Somar as vendas\r\n"
    'soma = 0\r\n'
    'contar = 0\r\n'
    'With GridNotas\r\n'
    '   For i = 1 To .rows - 1\r\n'
    '      If .TextMatrix(i, 7) = "Em Digita\xe7\xe3o" Then\r\n'
    "        'If .TextMatrix(i, 15) <> \"SIM\" Then\r\n"
    '            contar = contar + 1\r\n'
    '            soma = soma + CCur(.TextMatrix(i, 6))\r\n'
    "        'End If\r\n"
    '      End If\r\n'
    '   Next\r\n'
    'End With\r\n'
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
    'cmdEditar.Caption = "Editar"',
))

# ── Patch E: cmdImprimirConsulta_Click — ler do ListView em vez dos labels ────
patches.append(('imprimir_consulta',
    'REL_NFe_Consulta.dfQuant.Caption = lblQuantEnviada.Caption\r\n'
    'REL_NFe_Consulta.dfBruto.Caption = lblTotalEnviada.Caption',

    'REL_NFe_Consulta.dfQuant.Caption = lvwTotais.ListItems(1).SubItems(1)\r\n'
    'REL_NFe_Consulta.dfBruto.Caption = lvwTotais.ListItems(1).SubItems(2)',
))

# ── Verificar contagens (no texto apos Patch A) ───────────────────────────────
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
    print('\nOK — ListView substituiu Frame2 Totais')
else:
    print('\nABORTADO')
