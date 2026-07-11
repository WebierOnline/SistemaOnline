# -*- coding: utf-8 -*-
"""
Repara os 38 chameleonButton pre-existentes de OS_Recapadora.frm que
perderam suas propriedades (BTYPE/TX/cores/MICON/etc) quando o VB6
resalvou o arquivo apos o erro 380 do ChamaleonBtn. Restaura o bloco
COMPLETO de cada botao (por nome) a partir do checkpoint git 22bc784,
que tinha o arquivo intacto antes da extracao da aba CONSULTA.

cmdExibir e cmdImprimirConsulta NAO entram aqui - foram movidos para
OS_Consulta.frm e serao tratados (convertidos p/ CommandButton) em
outro script.
"""

CURRENT_PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"
CHECKPOINT_PATH = r"C:\Users\NOTEBOOK\AppData\Local\Temp\claude\C--projeto\916fb1c0-4fd5-437b-8d03-a83de36ec5b2\scratchpad\checkpoint_os_recapadora.frm"

with open(CURRENT_PATH, "rb") as f:
    cur_text = f.read().decode("cp1252")
cur_lines = cur_text.split("\r\n")

with open(CHECKPOINT_PATH, "rb") as f:
    chk_raw = f.read()
chk_text = chk_raw.decode("cp1252")
chk_lines = chk_text.split("\n")  # blob do git eh so LF


def find_block(lines, prefix, start=0):
    i = None
    for j in range(start, len(lines)):
        if lines[j].lstrip().startswith(prefix):
            i = j
            break
    if i is None:
        return None
    indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
    end_marker = indent + "End"
    for k in range(i + 1, len(lines)):
        if lines[k].rstrip("\r") == end_marker:
            return i, k
    return None


button_names = [
    "cmdCancelarParecer", "cmdSalvarParecer", "cmdCal2", "cmdCancelar",
    "cmdFinalizar", "cmdRemoverPecas", "cmdAdicionarPecas",
    "cmdRemoverServicosAuto", "cmdAdicionarServicosAuto",
    "cmdEditarServicosAuto", "ccmdIncluirSituacao", "cmdRemoverSituacao",
    "cmdAdicionarSituacao", "ccmdIncluirAcess", "cmdRemoverAcessorios",
    "cmdAdicionarAcessorios", "chameleonButton1", "cmdCal1",
    "cmdCancelarEntrada", "cmdAlterar", "cmdApagar", "cmdGerarEntrada",
    "cmdNovo", "cmdEditarOS", "cmdNovoOS", "cmdFinanceiroOS",
    "cmdImpEntrada2", "cmdImpOrcamento2", "cmdImpEntrada1",
    "cmdImpOrcamento1", "cmdImpPedido1", "cmdImpPedido2",
    "cmdImpGarantia1", "cmdOrcamentoPDF", "cmdPedidoPDF",
    "cmdFinalizarAV", "cmdFinalizarAP", "cmdExcluir",
]

assert len(button_names) == 38, len(button_names)

restaurados = 0
nao_encontrados_atual = []
for name in button_names:
    chk_block = find_block(chk_lines, f"Begin ChamaleonBtn.chameleonButton {name}")
    assert chk_block, f"nao achei {name} no checkpoint"
    cs, ce = chk_block
    chk_content = [l.rstrip("\r") for l in chk_lines[cs : ce + 1]]

    cur_block = find_block(cur_lines, f"Begin ChamaleonBtn.chameleonButton {name}")
    if cur_block is None:
        nao_encontrados_atual.append(name)
        continue
    us, ue = cur_block
    cur_lines[us : ue + 1] = chk_content
    restaurados += 1

print(f"restaurados {restaurados} botoes")
if nao_encontrados_atual:
    print("NAO encontrados no arquivo atual (nao restaurados):", nao_encontrados_atual)

out_text = "\r\n".join(cur_lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(CURRENT_PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - propriedades restauradas")
