import os

# ---- novo script (ASCII puro; 'i' acentuado via CHAR(237) pra nao depender de encoding no sqlcmd) ----
sql = (
"-- Preenche NotaFiscalItens.TipoProduto onde ficou NULL/'' - notas criadas por exe antigo que nao\n"
"-- gravava essa coluna. modNFe.bas (TransmitirNFe) faz  If NFeItens!TipoProduto = \"Combustivel\"  sem\n"
"-- tratar NULL, entao If Null Then -> erro 94 \"Invalid use of Null\" e a transmissao trava.\n"
"-- Mesma regra da procedure: COMBUSTIVEL=1 -> 'Combustivel' (com i acentuado), senao 'Produto'.\n"
"-- CHAR(237) = i com acento agudo em cp1252 (evita problema de encoding do arquivo no sqlcmd).\n"
"-- Idempotente.\n"
"IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('NotaFiscalItens') AND name = 'TipoProduto')\n"
"    ALTER TABLE NotaFiscalItens ADD TipoProduto varchar(20) NULL;\n"
"GO\n"
"\n"
"UPDATE ni\n"
"SET ni.TipoProduto = CASE WHEN ISNULL(p.COMBUSTIVEL, 0) = 1\n"
"                          THEN 'Combust' + CHAR(237) + 'vel'\n"
"                          ELSE 'Produto' END\n"
"FROM NotaFiscalItens ni\n"
"LEFT JOIN produtos p ON p.CODIGO = ni.CodigoProduto\n"
"WHERE ni.TipoProduto IS NULL OR LTRIM(RTRIM(ni.TipoProduto)) = '';\n"
"\n"
"SELECT COUNT(*) AS itens_ainda_sem_tipoproduto\n"
"FROM NotaFiscalItens WHERE TipoProduto IS NULL OR LTRIM(RTRIM(TipoProduto)) = '';\n"
)
b = sql.encode('ascii').replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
for base in (r'C:\Projeto\Arquivos\scripts', r'C:\scripts'):
    with open(os.path.join(base, 'Fix_NotaFiscalItens_TipoProduto.sql'), 'wb') as f:
        f.write(b)

# ---- manifesto: adiciona como 105 no fim (sem renumerar nada) ----
for base in (r'C:\Projeto\Arquivos\scripts', r'C:\scripts'):
    mpath = os.path.join(base, '_manifesto.txt')
    raw = open(mpath, 'rb').read().decode('cp1252')
    linhas = [l for l in raw.split('\n')]
    # descobre maior numero
    nums = []
    for l in linhas:
        s = l.strip().rstrip('\r')
        if '|' in s and s.split('|',1)[0].isdigit():
            nums.append(int(s.split('|',1)[0]))
    novo_num = max(nums) + 1
    # garante \n final e append
    txt = raw
    if not txt.endswith('\n'):
        txt += '\n'
    if 'Fix_NotaFiscalItens_TipoProduto.sql' not in txt:
        txt += '%03d|Fix_NotaFiscalItens_TipoProduto.sql|GERAL\n' % novo_num
    b2 = txt.encode('cp1252').replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
    open(mpath, 'wb').write(b2)

a = open(r'C:\Projeto\Arquivos\scripts\_manifesto.txt','rb').read()
c = open(r'C:\scripts\_manifesto.txt','rb').read()
print("manifesto identico:", a == c)
print("script identico:",
      open(r'C:\Projeto\Arquivos\scripts\Fix_NotaFiscalItens_TipoProduto.sql','rb').read()
      == open(r'C:\scripts\Fix_NotaFiscalItens_TipoProduto.sql','rb').read())
print("--- ultimas 4 linhas do manifesto ---")
for l in a.decode('cp1252').rstrip().split('\r\n')[-4:]:
    print(l)
print("--- script novo ---")
print(b.decode('ascii'))
