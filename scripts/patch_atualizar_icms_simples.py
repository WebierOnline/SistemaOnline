# -*- coding: utf-8 -*-
# Corrige AtualizarValorICMS em NFe_Completa.frm:
# Para emitentes Simples Nacional (RegimeTributario 1, 2 ou 5), zera vBC e
# vICMS de todos os itens e sai imediatamente, evitando que o UPDATE generico
# sobreponha vBC=ValorTotalBruto para itens com CSOSN (sem <vBC> no XML),
# o que causava "Total da BC ICMS difere do somatorio dos itens" no SEFAZ.
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = (
    'Sub AtualizarValorICMS()\r\n'
    'sSQL = "SELECT pRedBC as AliqRedBC FROM NotaFiscalItens'
)

NEW = (
    'Sub AtualizarValorICMS()\r\n'
    'Dim bSimplesAtuICMS As Boolean\r\n'
    'If vRegimeTributario > 0 Then\r\n'
    '    bSimplesAtuICMS = (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5)\r\n'
    'Else\r\n'
    '    Dim sTmpReg As String\r\n'
    '    sTmpReg = SQLExecutaRetorno("SELECT RegimeTributario FROM Empresa", "RegimeTributario")\r\n'
    '    Dim iTmpReg As Integer\r\n'
    '    iTmpReg = IIf(sTmpReg = "", 0, CInt(sTmpReg))\r\n'
    '    bSimplesAtuICMS = (iTmpReg = 1 Or iTmpReg = 2 Or iTmpReg = 5)\r\n'
    'End If\r\n'
    'If bSimplesAtuICMS Then\r\n'
    '    SQLExecuta "UPDATE NotaFiscalItens SET vBC = 0, vICMS = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    '    SQLExecuta "UPDATE NotaFiscal SET ValorICMS = 0, BaseICMS = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    '    txtValorICMS.Text = Format(0, ocMONEY)\r\n'
    '    txtBaseICMS.Text = Format(0, ocMONEY)\r\n'
    '    Exit Sub\r\n'
    'End If\r\n'
    'sSQL = "SELECT pRedBC as AliqRedBC FROM NotaFiscalItens'
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 1)')

if count == 1:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — AtualizarValorICMS agora respeita Simples Nacional: vBC=0')
else:
    print('ABORTADO')
