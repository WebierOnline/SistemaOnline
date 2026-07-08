# -*- coding: utf-8 -*-
"""
Patch OS_Recapadora.frm: atualiza referencias de codigo para os labels
que o usuario converteu em array (lblArray), removendo os controles
individuais antigos do .frm. So a secao de codigo precisa mudar - as
declaracoes Begin VB.Label ja foram substituidas manualmente pelo usuario.
"""
import re

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

MAPPING = {
    "lblMecanico": "lblArray(1)",
    "lblFabricante": "lblArray(7)",
    "lblModelo": "lblArray(8)",
    "lblAno": "lblArray(9)",
    "lblPlaca": "lblArray(10)",
    "lblKM": "lblArray(11)",
    "lblChassi": "lblArray(12)",
    "lblCor": "lblArray(13)",
    "lblTanque": "lblArray(14)",
}

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")

before_counts = {name: len(re.findall(r"\b" + name + r"\b", text)) for name in MAPPING}

for name, repl in MAPPING.items():
    text = re.sub(r"\b" + name + r"\b", repl, text)

after_counts = {name: len(re.findall(r"\b" + re.escape(name) + r"\b", text)) for name in MAPPING}
new_counts = {repl: text.count(repl) for repl in MAPPING.values()}

for name in MAPPING:
    assert after_counts[name] == 0, f"sobrou {name} sem substituir"

print("Ocorrencias substituidas:")
for name, repl in MAPPING.items():
    print(f"  {name} -> {repl}: {before_counts[name]}")

out = text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
print("bytes originais:", len(raw), "bytes finais:", len(out))
