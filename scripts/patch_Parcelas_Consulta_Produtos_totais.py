"""
patch_Parcelas_Consulta_Produtos_totais.py

Preenche lblTotalProdutos e lblTotalServico em Parcelas_Consulta_Produtos.frm.
As queries seguem o mesmo padrao DAO (dbData.OpenRecordset) ja usado no form.
"""
import sys, shutil

FRM = r"C:\Projeto\Compartilhado\Forms\Parcelas_Consulta_Produtos.frm"
BAK = FRM + ".bak_totais_prod_serv"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"If r.State <> 0 Then r.Close\r\n"
    b"Set r = Nothing\r\n"
    b"\r\n"
    b"txtCodPedido.Text = Format(Pedido, \"000000\")\r\n"
    b"End Sub\r\n"
)

NEW = (
    b"If r.State <> 0 Then r.Close\r\n"
    b"Set r = Nothing\r\n"
    b"\r\n"
    b"sSQL = \"SELECT ISNULL(SUM(total),0) AS vProd FROM pedidos_itens WHERE cod_pedido=\" & Pedido\r\n"
    b"Set r = dbData.OpenRecordset(sSQL)\r\n"
    b"lblTotalProdutos.Caption = Format(r(\"vProd\"), ocMONEY)\r\n"
    b"If r.State <> 0 Then r.Close\r\n"
    b"Set r = Nothing\r\n"
    b"\r\n"
    b"sSQL = \"SELECT ISNULL(SUM(sv.total),0) AS vServ FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=\" & Pedido\r\n"
    b"Set r = dbData.OpenRecordset(sSQL)\r\n"
    b"lblTotalServico.Caption = Format(r(\"vServ\"), ocMONEY)\r\n"
    b"If r.State <> 0 Then r.Close\r\n"
    b"Set r = Nothing\r\n"
    b"\r\n"
    b"txtCodPedido.Text = Format(Pedido, \"000000\")\r\n"
    b"End Sub\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: lblTotalProdutos e lblTotalServico preenchidos em loadPedidos")
print("Arquivo salvo com sucesso.")
