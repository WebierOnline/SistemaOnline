"""
patch_cmdrecalcular.py
Ativa cmdRecalcular_Click chamando RecalcularItensNota.

RecalcularItensNota ja faz tudo:
  - Percorre NotaFiscalItens e recalcula ICMS, PIS, COFINS, IPI, ICMS-ST,
    IBS/CBS e IS de cada item (UPDATE filho).
  - Ao final chama Exibir_Itens e AtualizarTotaisNota (UPDATE pai).
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_cmdrecalcular"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"Private Sub cmdRecalcular_Click()\r\n"
    b"'Call MostrarValorItens  'desativei pq estava calculando o valortotalbruto errado\r\n"
    b"'Call AtualizarValorICMS\r\n"
    b"'If Left(cboDestOperacao.Text, 1) = \"2\" Then Call CalcularICMSInterItens\r\n"
    b"'Call DistribuirFrete\r\n"
    b"'Call DistribuirSeguro\r\n"
    b"'Call DistribuirOutros\r\n"
    b"'Call DistribuirDesconto\r\n"
    b"'Call AtualizarTotaisNota\r\n"
    b"End Sub\r\n"
)

NEW = (
    b"Private Sub cmdRecalcular_Click()\r\n"
    b"    If txtCodNota.Text = \"\" Then Exit Sub\r\n"
    b"    RecalcularItensNota\r\n"
    b"End Sub\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: count={cnt} (esperado 1)")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)
print("OK   P1")
print("\nArquivo salvo com sucesso.")
