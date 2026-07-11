# -*- coding: utf-8 -*-
"""
Reduz o SSTab1 de 6 para 5 abas, removendo a aba CONSULTA (indice 3) e
renumerando as abas 4 (vazia) e 5 (lblQuantFiltro) para 3 e 4.
Tambem corrige SSTab1.TabVisible(5) -> TabVisible(4) (mesmo bug
encontrado na 1a limpeza - a referencia de indice de aba precisa
acompanhar a renumeracao).
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

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


i_tabs = find_line_exact("      Tabs            =   6")
i_tabsperrow = find_line_exact("      TabsPerRow      =   6")

i_block_start = find_line_exact('      TabCaption(3)   =   "CONSULTA"')
i_block_end = find_line_exact('      Tab(5).ControlCount=   1')

old_block = lines[i_block_start : i_block_end + 1]
expected_old_block = [
    '      TabCaption(3)   =   "CONSULTA"',
    '      TabPicture(3)   =   "OS_Recapadora.frx":2495',
    "      Tab(3).ControlEnabled=   0   'False",
    '      Tab(3).Control(0)=   "Frame2"',
    "      Tab(3).Control(0).Enabled=   0   'False",
    '      Tab(3).Control(1)=   "Grid"',
    "      Tab(3).Control(1).Enabled=   0   'False",
    '      Tab(3).Control(2)=   "lblQuant"',
    "      Tab(3).Control(2).Enabled=   0   'False",
    '      Tab(3).Control(3)=   "lblTotalConsulta"',
    "      Tab(3).Control(3).Enabled=   0   'False",
    "      Tab(3).ControlCount=   4",
    '      TabCaption(4)   =   " "',
    '      TabPicture(4)   =   "OS_Recapadora.frx":24B1',
    "      Tab(4).ControlEnabled=   0   'False",
    "      Tab(4).ControlCount=   0",
    '      TabCaption(5)   =   " "',
    '      TabPicture(5)   =   "OS_Recapadora.frx":24CD',
    "      Tab(5).ControlEnabled=   0   'False",
    '      Tab(5).Control(0)=   "lblQuantFiltro"',
    "      Tab(5).ControlCount=   1",
]
assert old_block == expected_old_block, old_block

new_block = [
    '      TabCaption(3)   =   " "',
    '      TabPicture(3)   =   "OS_Recapadora.frx":24B1',
    "      Tab(3).ControlEnabled=   0   'False",
    "      Tab(3).ControlCount=   0",
    '      TabCaption(4)   =   " "',
    '      TabPicture(4)   =   "OS_Recapadora.frx":24CD',
    "      Tab(4).ControlEnabled=   0   'False",
    '      Tab(4).Control(0)=   "lblQuantFiltro"',
    "      Tab(4).ControlCount=   1",
]

lines[i_tabs] = "      Tabs            =   5"
lines[i_tabsperrow] = "      TabsPerRow      =   5"
lines[i_block_start : i_block_end + 1] = new_block

# corrige a referencia de indice de aba (TabVisible) que nao acompanha
# a renumeracao automaticamente
i_tv = find_line_exact("SSTab1.TabVisible(5) = False")
lines[i_tv] = "SSTab1.TabVisible(4) = False"

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - SSTab1 renumerado para 5 abas + TabVisible corrigido")
