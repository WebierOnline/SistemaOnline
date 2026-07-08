# -*- coding: utf-8 -*-
"""
Corrige as 13 linhas que o repair_encoding_corruption.py nao conseguiu
alinhar automaticamente com o HEAD (sao linhas novas que eu mesmo escrevi
nesta sessao, ou legendas adicionadas pelo usuario apos o HEAD baseline).
Digitando o texto correto diretamente aqui (verificado sem nenhum
caractere de substituicao antes de rodar).
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")

TARGET = chr(0xEF) + chr(0xBF) + chr(0xBD)

FIXES = {
    1: 'Caption         =   "Mecânico"',
    2: 'If vTipoOS <> "Automóveis" And vTipoOS <> "Motocicletas" And vTipoOS <> "Informática" And vTipoOS <> "Celular" Then Exit Sub',
    3: 'If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 2) <> "SERVIÇO" Then Exit Sub',
    4: 'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then',
    5: 'If vCodMecanicoServ = "" Then MsgBox "Selecione o mecânico que executou o serviço!", vbExclamation, "Aviso do Sistema": Exit Sub',
    6: "'CHECAR SE A OS ESTÁ FECHADA",
    7: 'ElseIf vTipoOS = "Comunicação Visual" Then',
}

# localiza cada linha alvo pelo trecho ainda intacto (prefixo/sufixo sem a corrupcao)
targets = [
    ("Caption         =   \"Mec", FIXES[1]),
    ('If vTipoOS <> "Autom', FIXES[2]),
    ('If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 2) <>', FIXES[3]),
]

count_fixed = 0
for i, line in enumerate(lines):
    if TARGET not in line:
        continue
    if line.startswith("               Begin VB.Label lblMecanicoServ"):
        continue
    if 'Caption         =   "Mec' in line and TARGET in line:
        lines[i] = FIXES[1]
        count_fixed += 1
    elif line.startswith('If vTipoOS <> "Autom'):
        lines[i] = FIXES[2]
        count_fixed += 1
    elif 'Grid_Servicos.TextMatrix(Grid_Servicos.Row, 2) <>' in line:
        lines[i] = FIXES[3]
        count_fixed += 1
    elif line.strip().startswith('If vTipoOS = "Autom') and 'Celular' in line and line.strip().startswith('If '):
        lines[i] = ("        " if line.startswith("        ") else "") + FIXES[4]
        count_fixed += 1
    elif 'Selecione o mec' in line and 'MsgBox' in line:
        lines[i] = FIXES[5]
        count_fixed += 1
    elif line.strip() == "'CHECAR SE A OS EST" + TARGET + " FECHADA":
        lines[i] = FIXES[6]
        count_fixed += 1
    elif line.strip().startswith('If vTipoOS = "Autom') and line.strip().endswith('Then') and 'Or vTipoOS = "Motocicletas" Then' in line:
        indent = line[: len(line) - len(line.lstrip())]
        lines[i] = indent + 'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then'
        count_fixed += 1
    elif line.strip().startswith('ElseIf vTipoOS = "Inform') and 'Celular' in line:
        indent = line[: len(line) - len(line.lstrip())]
        lines[i] = indent + 'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then'
        count_fixed += 1
    elif line.strip().startswith('ElseIf vTipoOS = "Comunica'):
        indent = line[: len(line) - len(line.lstrip())]
        lines[i] = indent + FIXES[7]
        count_fixed += 1

print("linhas corrigidas nesta passada:", count_fixed)

remaining = [i + 1 for i, l in enumerate(lines) if TARGET in l]
print("restantes com corrupcao:", len(remaining))
for r in remaining:
    print(" ", r, repr(lines[r - 1][:150]))

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")

with open(PATH, "wb") as f:
    f.write(out)

print("saved, bytes:", len(out))
