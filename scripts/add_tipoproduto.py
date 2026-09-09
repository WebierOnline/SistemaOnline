import os

# ---------- 1) novo script ----------
sql = (
"-- Adiciona TbNFCe_Itens.TipoProduto - coluna usada pelas procedures NFCeDuplicar e NFCeIncluir\n"
"-- ('Produto' / 'Combustivel', vindo de CASE WHEN produtos.COMBUSTIVEL = 1). Existe nos bancos\n"
"-- antigos por schema acumulado, mas nao era criada por nenhum script do manifesto - cliente com\n"
"-- banco-base mais novo dava 'Nome de coluna TipoProduto invalido' (Msg 207) no script NFCeDuplicar.\n"
"-- NULL sem DEFAULT = metadata-only em qualquer versao do SQL Server (inclusive 2008 Express).\n"
"IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'TipoProduto')\n"
"    ALTER TABLE TbNFCe_Itens ADD TipoProduto varchar(20) NULL;\n"
)
b = sql.encode('cp1252').replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
for base in (r'C:\Projeto\Arquivos\scripts', r'C:\scripts'):
    with open(os.path.join(base, 'AlterTable_TbNFCe_Itens_TipoProduto.sql'), 'wb') as f:
        f.write(b)

# ---------- 2) manifesto: insere antes do 096 e renumera dali pra frente ----------
NOVO = 'AlterTable_TbNFCe_Itens_TipoProduto.sql'
for base in (r'C:\Projeto\Arquivos\scripts', r'C:\scripts'):
    mpath = os.path.join(base, '_manifesto.txt')
    raw = open(mpath, 'rb').read().decode('cp1252')
    linhas = raw.split('\n')
    out = []
    inserido = False
    for ln in linhas:
        s = ln.rstrip('\r')
        if not s.strip():
            out.append(s)
            continue
        num_str, resto = s.split('|', 1)
        try:
            num = int(num_str)
        except ValueError:
            out.append(s)
            continue
        if num == 96 and not inserido:
            out.append('096|%s|GERAL' % NOVO)
            inserido = True
        if num >= 96:
            out.append('%03d|%s' % (num + 1, resto))
        else:
            out.append(s)
    novo_raw = '\n'.join(out)
    b2 = novo_raw.encode('cp1252').replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
    open(mpath, 'wb').write(b2)

# ---------- 3) confere ----------
import subprocess
a = open(r'C:\Projeto\Arquivos\scripts\_manifesto.txt','rb').read()
c = open(r'C:\scripts\_manifesto.txt','rb').read()
print("manifesto identico nas 2 pastas:", a == c)
print("script identico nas 2 pastas:",
      open(r'C:\Projeto\Arquivos\scripts\AlterTable_TbNFCe_Itens_TipoProduto.sql','rb').read()
      == open(r'C:\scripts\AlterTable_TbNFCe_Itens_TipoProduto.sql','rb').read())
print("--- trecho do manifesto 092-106 ---")
for ln in a.decode('cp1252').split('\r\n'):
    if ln[:3].isdigit() and 92 <= int(ln[:3]) <= 106:
        print(ln)
