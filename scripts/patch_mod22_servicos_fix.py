# -*- coding: utf-8 -*-
"""
Fix FormatarGrid_Servicos em Vendas_Consulta_PorProdutos.frm:
  1. sBase SQL: adiciona s.cod_servico AS varCodServ
  2. ColWidth(12) = 0 (ocultar coluna CÓD. PRODUTO para servicos)
  3. Header col4: "CÓD. BARRA" -> "CÓD. SERV."
  4. Data col4: "" -> rTabela("varCodServ")
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# 1. sBase SQL: adiciona s.cod_servico AS varCodServ apos var_CodOS
old_sbase = (
    "s.desconto AS varDesc, s.total AS varTotal, ISNULL(OS.COD_PEDIDO, 0) AS var_CodOS \" & _\n"
    "           \"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS \"\n"
)
new_sbase = (
    "s.desconto AS varDesc, s.total AS varTotal, ISNULL(OS.COD_PEDIDO, 0) AS var_CodOS, s.cod_servico AS varCodServ \" & _\n"
    "           \"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS \"\n"
)
changes.append((old_sbase, new_sbase, 'sBase SQL — adiciona s.cod_servico AS varCodServ', False))

# 2. ColWidth(12) = 900 -> 0 em FormatarGrid_Servicos
# Ancora unica: TextMatrix(0, 1) = "OS" distingue de ProdDetalhado (que tem "PEDIDO")
old_colw = (
    "      .ColWidth(11) = 0\n"
    "      .ColWidth(12) = 900\n"
    "      \n"
    "      .TextMatrix(0, 1) = \"OS\"\n"
)
new_colw = (
    "      .ColWidth(11) = 0\n"
    "      .ColWidth(12) = 0\n"
    "      \n"
    "      .TextMatrix(0, 1) = \"OS\"\n"
)
changes.append((old_colw, new_colw, 'FormatarGrid_Servicos ColWidth(12) = 0 (ocultar)', False))

# 3. Header col4: "CÓD. BARRA" -> "CÓD. SERV." em FormatarGrid_Servicos
# Ancora unica: TextMatrix(0, 3) = "" e' unico de Servicos (ProdDetalhado tem "CÓD.PROD.")
old_hdr4 = (
    "      .TextMatrix(0, 3) = \"\"\n"
    "      .TextMatrix(0, 4) = \"CÓD. BARRA\"\n"
)
new_hdr4 = (
    "      .TextMatrix(0, 3) = \"\"\n"
    "      .TextMatrix(0, 4) = \"CÓD. SERV.\"\n"
)
changes.append((old_hdr4, new_hdr4, 'FormatarGrid_Servicos header col4 — CÓD. SERV.', False))

# 4. Data col4: "" -> rTabela("varCodServ") em FormatarGrid_Servicos
# Ancora unica: TextMatrix(.rows-1, 3) = "" seguido de col4 = "" (unico de Servicos)
old_data4 = (
    "            .TextMatrix(.rows - 1, 3) = \"\"\n"
    "            .TextMatrix(.rows - 1, 4) = \"\"\n"
)
new_data4 = (
    "            .TextMatrix(.rows - 1, 3) = \"\"\n"
    "            .TextMatrix(.rows - 1, 4) = rTabela(\"varCodServ\")\n"
)
changes.append((old_data4, new_data4, 'FormatarGrid_Servicos data col4 — rTabela("varCodServ")', False))

for old, new, label, replace_all in changes:
    count = text.count(old)
    if replace_all:
        if count == 0:
            print(f'ERRO [{label}]: 0 ocorrencias')
            sys.exit(1)
    else:
        if count != 1:
            print(f'ERRO [{label}]: {count} ocorrencias (esperado 1)')
            sys.exit(1)
    text = text.replace(old, new)
    print(f'OK ({count}x): {label}')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
