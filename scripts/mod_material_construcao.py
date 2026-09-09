import sys

def patch(path, old, new, occ=1):
    b = open(path, 'rb').read()
    n = b.count(old)
    if n != occ:
        print("ABORT %s: ancora aparece %d vezes (esperado %d)" % (path, n, occ))
        print("  ancora:", old[:80])
        sys.exit(1)
    b = b.replace(old, new, occ)
    b = b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
    open(path, 'wb').write(b)
    print("OK", path)

R = r'C:\Projeto'

# ---------- PARTE 1: Tela_Principal.frm - Menu_PROD para tipos 6/7/8 ----------
patch(R + r'\OnlineCommerce\Forms\Tela_Principal.frm',
      b'   If varTipoEmpresa < "6" Then\r\n',
      b'   \x27Tipos 1..8 (todos varejo em Configuracao_Geral) mostram produtos. Os ElseIf de "6"/"7"/"8"\r\n'
      b'   \x27abaixo viraram codigo morto (nao ha opcao Escola/Curso no sistema) - mantidos so por historico.\r\n'
      b'   If Val(varTipoEmpresa) >= 1 And Val(varTipoEmpresa) <= 8 Then\r\n')

# ---------- PARTE 2a: unidade "GL" nas 3 listas ----------
patch(R + r'\Compartilhado\Forms\Produtos_Cadastro.frm',
      b'        .AddItem "DI": .AddItem "FD": .AddItem "PT"\r\n    End With',
      b'        .AddItem "DI": .AddItem "FD": .AddItem "PT": .AddItem "GL"\r\n    End With')

patch(R + r'\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm',
      b'         .AddItem "DI": .AddItem "FD": .AddItem "PT"\r\n      End With',
      b'         .AddItem "DI": .AddItem "FD": .AddItem "PT": .AddItem "GL"\r\n      End With')

patch(R + r'\Compartilhado\Forms\Produtos_CadastoRapido.frm',
      b'   cboUnidMedida.AddItem "GR"\r\n   moCombo.AttachTo cboUnidMedida',
      b'   cboUnidMedida.AddItem "GR"\r\n   cboUnidMedida.AddItem "GL"\r\n   moCombo.AttachTo cboUnidMedida')

# ---------- PARTE 2b: tira -f 65001 do runner (scripts sao cp1252) ----------
patch(R + r'\Arquivos\scripts\Executar_Scripts.bat',
      b'-i "%SCRIPTDIR%!ARQ!" -f 65001 -b',
      b'-i "%SCRIPTDIR%!ARQ!" -b')

patch(R + r'\Arquivos\scripts\Executar_Scripts.ps1',
      b'-i "$caminhoScript" -f 65001 -b',
      b'-i "$caminhoScript" -b')

# espelha runners em C:\scripts
import shutil
for f in ('Executar_Scripts.bat', 'Executar_Scripts.ps1'):
    shutil.copyfile(R + r'\Arquivos\scripts\\' + f, r'C:\scripts\\' + f)
    a = open(R + r'\Arquivos\scripts\\' + f, 'rb').read()
    c = open(r'C:\scripts\\' + f, 'rb').read()
    print("mirror", f, "identico:", a == c)
