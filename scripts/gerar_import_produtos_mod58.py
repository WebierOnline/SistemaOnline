# -*- coding: utf-8 -*-
"""mod58: gera o INSERT de importacao da planilha Produtos.csv (CODIGO, DESCRICAO, CUSTO,
VALOR_VV) pras tabelas produtos e Produtos_Precos de UM cliente especifico (nao e script de
deploy pra todo mundo - nao entra no manifesto nem espelha pra C:\\scripts).

- CSV e UTF-8 de verdade (Ç/Ã/Á/É corretos) - o "acento errado" que o usuario viu (imagem
  image259) e o Excel abrindo um CSV UTF-8 como ANSI (mojibake so na visualizacao dele).
  Le com csv.reader direto, sem "corrigir" nada.
- CUSTO/VALOR_VV do CSV ja vem com ponto decimal ("7.69") = ja e literal T-SQL valido, nao
  precisa de nenhuma conversao ponto<->virgula (a virgula que aparece no banco/telas do VB6
  e so formatacao de exibicao pt-BR, nao afeta o literal da INSERT).
- Campos auto-preenchidos em produtos e o padrao de Produtos_Precos (Codigo por MAX()+1,
  COD_ENTRADA = proprio codigo do produto, FORMA = 'CADASTRO', Data = hoje sem separador)
  replicam exatamente o que Produtos_Cadastro.Preco_Entrada ja faz ao cadastrar um produto
  novo manualmente (Compartilhado/Forms/Produtos_Cadastro.frm ~linha 4362).
"""
import csv
import datetime

CSV_PATH = r"C:\Projeto\Arquivos\Produtos.csv"
OUT_PATH = r"C:\Projeto\Arquivos\Import_Produtos_mod58.sql"

with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
    rdr = csv.reader(f, delimiter=";")
    header = next(rdr)
    rows = [r for r in rdr if any(c.strip() for c in r)]

def sqlstr(s):
    return "'" + s.replace("'", "''") + "'"

ok_rows = []
bad_rows = []
custo_zerado = []  # linhas onde o usuario pediu pra usar 0,00 no CUSTO (linha 273/702 do CSV - "C3"/"MOSTRUARIO")
for i, r in enumerate(rows, 2):  # linha 2 = primeira linha de dado (1 e o cabecalho)
    cod, desc, custo, vv = r[0].strip(), r[1].strip(), r[2].strip(), r[3].strip()
    problema = None
    custo_f = None
    if not cod.isdigit():
        problema = "codigo invalido: %r" % cod
    elif desc == "":
        problema = "descricao vazia"
    else:
        try:
            custo_f = float(custo)
        except ValueError:
            custo_f = 0.00   # pedido do usuario: usar 0,00 quando o CUSTO da planilha nao e numero
            custo_zerado.append((i, cod, desc, custo))
        try:
            vv_f = float(vv)
        except ValueError:
            problema = "VALOR_VV nao numerico: %r" % vv
    if problema:
        bad_rows.append((i, cod, desc, custo, vv, problema))
    else:
        ok_rows.append((int(cod), desc, custo_f, vv_f))

hoje = datetime.date.today().strftime("%Y%m%d")

linhas = []
linhas.append("-- mod58: importacao de %d produtos (Arquivos/Produtos.csv) num cliente especifico." % len(ok_rows))
linhas.append("-- NAO e script de deploy (nao vai pro manifesto/C:\\scripts) - rodar manualmente so nesse banco.")
linhas.append("-- %d linha(s) com problema NAO foram incluidas (ver relatorio no console do gerador)." % len(bad_rows))
if custo_zerado:
    linhas.append("-- %d linha(s) com CUSTO invalido na planilha (texto no lugar de numero) -> CUSTO = 0.00 (pedido do usuario):" % len(custo_zerado))
    for z in custo_zerado:
        linhas.append("--   codigo %s: %s (custo original na planilha: %r)" % (z[1], z[2], z[3]))
linhas.append("")
linhas.append("BEGIN TRANSACTION;")
linhas.append("")
linhas.append("-- ===================== produtos =====================")
for cod, desc, custo_f, vv_f in ok_rows:
    linhas.append(
        "INSERT INTO produtos (codigo, descricao, quant_estoque, unid_medida, ativo, cfop, ICMSCST, ICMSAliq, "
        "PISCST, COFINSCST, IPICST, NCM, CEST, EAN, Alterado, PedirPeso, PISAliq, COFINSAliq, IPIAliq, "
        "USOCONSUMO, COMBUSTIVEL, MATERIAPRIMA, IMOBILIZADO, FRACIONADO) VALUES (" +
        "%d, %s, 1, 'UN', 1, 5102, 102, '0.00', '04', '04', 99, '00000000', '0', 'SEM GTIN', 0, 0, "
        "'0.00', '0.00', '0.00', 0, 0, 0, 0, 0);" % (cod, sqlstr(desc))
    )

linhas.append("")
linhas.append("-- ===================== Produtos_Precos =====================")
linhas.append("DECLARE @precoCodigo INT;")
linhas.append("SELECT @precoCodigo = ISNULL(MAX(codigo), 0) FROM Produtos_Precos;")
linhas.append("")
for cod, desc, custo_f, vv_f in ok_rows:
    linhas.append("SET @precoCodigo = @precoCodigo + 1;")
    linhas.append(
        "INSERT INTO Produtos_Precos (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, CUSTO, "
        "MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP) VALUES (" +
        "@precoCodigo, %d, CONVERT(DATETIME, '%s'), %d, 'CADASTRO', %.2f, "
        "0.00, %.2f, 0.00, %.2f, 0.00, %.2f, 0.00, %.2f);"
        % (cod, hoje, cod, custo_f, vv_f, vv_f, vv_f, vv_f)
    )

linhas.append("")
linhas.append("COMMIT TRANSACTION;")

texto = "\r\n".join(linhas) + "\r\n"
open(OUT_PATH, "wb").write(texto.encode("cp1252"))  # .sql do projeto e sempre cp1252 (padrao dos outros scripts)

print("gerado:", OUT_PATH)
print("linhas OK:", len(ok_rows))
print("linhas com CUSTO zerado (pedido do usuario):", len(custo_zerado))
for z in custo_zerado:
    print("  linha %d: codigo=%s desc=%r custo original=%r -> 0.00" % z)
print("linhas com problema (nao incluidas):", len(bad_rows))
for b in bad_rows:
    print("  linha %d: codigo=%s desc=%r custo=%r valor_vv=%r -- %s" % b)
