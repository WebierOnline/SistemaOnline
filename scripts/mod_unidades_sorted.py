import sys

def patch(path, old, new):
    b = open(path, 'rb').read()
    n = b.count(old)
    if n != 1:
        print("ABORT %s: ancora aparece %d vezes" % (path, n)); sys.exit(1)
    b = b.replace(old, new, 1)
    b = b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
    open(path, 'wb').write(b)
    print("OK", path)

R = r'C:\Projeto'

# ---- Produtos_Cadastro.frm : lista completa (19 orig + GL) em ordem alfabetica ----
patch(R + r'\Compartilhado\Forms\Produtos_Cadastro.frm',
      b'    With cboUnidMedida\r\n'
      b'        .Clear\r\n'
      b'        .AddItem "UN": .AddItem "CX": .AddItem "M": .AddItem "M2"\r\n'
      b'        .AddItem "M3": .AddItem "ML": .AddItem "KG": .AddItem "GR"\r\n'
      b'        .AddItem "CT": .AddItem "PO": .AddItem "SC": .AddItem "PA"\r\n'
      b'        .AddItem "EX": .AddItem "BJ": .AddItem "DZ": .AddItem "PC"\r\n'
      b'        .AddItem "DI": .AddItem "FD": .AddItem "PT"\r\n'
      b'    End With',
      b'    With cboUnidMedida\r\n'
      b'        .Clear\r\n'
      b'        .AddItem "BJ": .AddItem "CT": .AddItem "CX": .AddItem "DI"\r\n'
      b'        .AddItem "DZ": .AddItem "EX": .AddItem "FD": .AddItem "GL"\r\n'
      b'        .AddItem "GR": .AddItem "KG": .AddItem "M":  .AddItem "M2"\r\n'
      b'        .AddItem "M3": .AddItem "ML": .AddItem "PA": .AddItem "PC"\r\n'
      b'        .AddItem "PO": .AddItem "PT": .AddItem "SC": .AddItem "UN"\r\n'
      b'    End With')

# ---- Produtos_Estoque_Simples.frm : mesma lista (controle cboEdit) ----
patch(R + r'\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm',
      b'      With cboEdit\r\n'
      b'         .AddItem "UN": .AddItem "CX": .AddItem "M":   .AddItem "M2"\r\n'
      b'         .AddItem "M3": .AddItem "ML": .AddItem "KG":  .AddItem "GR"\r\n'
      b'         .AddItem "CT": .AddItem "PO": .AddItem "SC":  .AddItem "PA"\r\n'
      b'         .AddItem "EX": .AddItem "BJ": .AddItem "DZ":  .AddItem "PC"\r\n'
      b'         .AddItem "DI": .AddItem "FD": .AddItem "PT"\r\n'
      b'      End With',
      b'      With cboEdit\r\n'
      b'         .AddItem "BJ": .AddItem "CT": .AddItem "CX": .AddItem "DI"\r\n'
      b'         .AddItem "DZ": .AddItem "EX": .AddItem "FD": .AddItem "GL"\r\n'
      b'         .AddItem "GR": .AddItem "KG": .AddItem "M":  .AddItem "M2"\r\n'
      b'         .AddItem "M3": .AddItem "ML": .AddItem "PA": .AddItem "PC"\r\n'
      b'         .AddItem "PO": .AddItem "PT": .AddItem "SC": .AddItem "UN"\r\n'
      b'      End With')

# ---- Produtos_CadastoRapido.frm : lista curta (8 orig + GL) em ordem alfabetica ----
patch(R + r'\Compartilhado\Forms\Produtos_CadastoRapido.frm',
      b'   cboUnidMedida.Clear\r\n'
      b'   cboUnidMedida.AddItem "UN"\r\n'
      b'   cboUnidMedida.AddItem "CX"\r\n'
      b'   cboUnidMedida.AddItem "M"\r\n'
      b'   cboUnidMedida.AddItem "M\xb2"\r\n'
      b'   cboUnidMedida.AddItem "M\xb3"\r\n'
      b'   cboUnidMedida.AddItem "ML"\r\n'
      b'   cboUnidMedida.AddItem "KG"\r\n'
      b'   cboUnidMedida.AddItem "GR"\r\n',
      b'   cboUnidMedida.Clear\r\n'
      b'   cboUnidMedida.AddItem "CX"\r\n'
      b'   cboUnidMedida.AddItem "GL"\r\n'
      b'   cboUnidMedida.AddItem "GR"\r\n'
      b'   cboUnidMedida.AddItem "KG"\r\n'
      b'   cboUnidMedida.AddItem "M"\r\n'
      b'   cboUnidMedida.AddItem "M\xb2"\r\n'
      b'   cboUnidMedida.AddItem "M\xb3"\r\n'
      b'   cboUnidMedida.AddItem "ML"\r\n'
      b'   cboUnidMedida.AddItem "UN"\r\n')
