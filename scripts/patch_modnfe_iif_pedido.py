p = r'C:\Projeto\Arquivos\scripts'  # not used; real target below
TARGET = r'C:\Projeto\modNFe_path.txt'

import io, sys

path = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
d = open(path, 'rb').read()

# ---- Replacement A: insere bloco de variaveis antes da chamada GerarItens ----
ancA = (b"        infAdiProd = Trim(infAdiProd)\r\n"
        b"        \r\n"
        b"        iRetorno = sistNFe.GerarItens(i, NFeItens!CodigoProduto,")
bloco = (b"        infAdiProd = Trim(infAdiProd)\r\n"
         b"        \r\n"
         b"        ' Cod_Pedido / Item_pedido podem vir NULL em notas antigas. IIf no VB6 avalia os\r\n"
         b"        ' dois lados sempre, entao CLng(NULL)/CStr(NULL) dava \"Uso invalido de Null\" (erro 94).\r\n"
         b"        Dim lItemPed As Long, sPedido As String\r\n"
         b"        If IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0 Then\r\n"
         b"            lItemPed = 0\r\n"
         b"            sPedido = \"\"\r\n"
         b"        Else\r\n"
         b"            sPedido = CStr(NFeItens!Cod_Pedido)\r\n"
         b"            If IsNull(NFeItens!Item_pedido) Then\r\n"
         b"                lItemPed = 0\r\n"
         b"            Else\r\n"
         b"                lItemPed = CLng(NFeItens!Item_pedido)\r\n"
         b"            End If\r\n"
         b"        End If\r\n"
         b"        \r\n"
         b"        iRetorno = sistNFe.GerarItens(i, NFeItens!CodigoProduto,")

# ---- Replacement B: troca os dois IIf pelos valores ja calculados ----
ancB = (b'IIf(IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0, 0, CLng(NFeItens!Item_pedido)), '
        b'IIf(IsNull(NFeItens!Cod_Pedido) Or NFeItens!Cod_Pedido = 0, "", CStr(NFeItens!Cod_Pedido))')
newB = b'lItemPed, sPedido'

for label, anc in (('A', ancA), ('B', ancB)):
    n = d.count(anc)
    if n != 1:
        print('ABORTOU: ancora %s aparece %d vezes (esperado 1)' % (label, n))
        sys.exit(1)

d = d.replace(ancA, bloco, 1)
d = d.replace(ancB, newB, 1)

# normaliza CRLF
d = d.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')

open(path, 'wb').write(d)
print('OK - patch aplicado em', path)

# mostra o trecho resultante
i = d.find(b'Cod_Pedido / Item_pedido podem vir NULL')
print('----')
print(d[i-40:i+950].decode('cp1252'))
