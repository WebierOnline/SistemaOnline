# -*- coding: utf-8 -*-
# Implementa campo uTrib_IS no fluxo NF-e IS.
#
# Alteracoes em NFe_Completa.frm (Load_Data_Itens):
#   A - Declaracao de variavel vISuTrib
#   B - Inicializacao antes do SELECT tbISClassTrib
#   C1/C2 - Limpeza em 2 blocos de reset
#   D - Adiciona uTrib_IS ao SELECT de tbISClassTrib
#   E - Le uTrib_IS do recordset rIS
#   F - Popula Tb("uTrib_IS")
#
# Alteracao em modNFe.bas (GerarItensImpostoIS):
#   G - Usa NFeItens!uTrib_IS em vez de NFeItens!UnidadeTributavel (2 ocorrencias)

import sys

# ══════════════════════════════════════════════════════════════════
# NFe_Completa.frm
# ══════════════════════════════════════════════════════════════════
PATH_FRM = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
text = open(PATH_FRM, 'rb').read().decode('windows-1252')
ok = True

patches = [
    # A - Declaracao de variavel
    (
        'A - Dim vISuTrib declaration',
        1,
        'Dim vISFatorConv As String      \'bloco IS\r\n',
        'Dim vISFatorConv As String      \'bloco IS\r\nDim vISuTrib As String          \'bloco IS\r\n',
    ),
    # B - Inicializar antes do SELECT (valor "0,0000")
    (
        'B - init vISuTrib before SELECT',
        1,
        '     vISvUnid = "0,0000"\r\n     Dim sClassIS As String',
        '     vISvUnid = "0,0000"\r\n     vISuTrib = ""\r\n     Dim sClassIS As String',
    ),
    # C1 - Reset no bloco sem indentacao (fin de sub antes de Load_Data_Duplicatas)
    (
        'C1 - reset vISuTrib (no-indent End Sub)',
        1,
        'vISvUnid = ""\r\nvISFatorConv = ""\r\nEnd Sub\r\n\r\nSub Load_Data_Duplicatas()',
        'vISvUnid = ""\r\nvISuTrib = ""\r\nvISFatorConv = ""\r\nEnd Sub\r\n\r\nSub Load_Data_Duplicatas()',
    ),
    # C2 - Reset no bloco com 5 espacos (dentro de If ... End If)
    (
        'C2 - reset vISuTrib (5-space End If)',
        1,
        '     vISvUnid = ""\r\n     vISFatorConv = ""\r\n End If\r\n If r.State <> 0 Then r.Close\r\nEnd Sub',
        '     vISvUnid = ""\r\n     vISuTrib = ""\r\n     vISFatorConv = ""\r\n End If\r\n If r.State <> 0 Then r.Close\r\nEnd Sub',
    ),
    # D - Adicionar uTrib_IS ao SELECT
    (
        'D - add uTrib_IS to SELECT',
        1,
        '"SELECT TOP 1 tipo_calculo_is, ISpAliq, ISvUnid " & _\r\n',
        '"SELECT TOP 1 tipo_calculo_is, ISpAliq, ISvUnid, uTrib_IS " & _\r\n',
    ),
    # E - Ler uTrib_IS do rIS
    (
        'E - read uTrib_IS from rIS',
        1,
        '           vISvUnid = Format(ValidateNull(rIS("ISvUnid")), "##,##0.0000")\r\n        End If',
        '           vISvUnid = Format(ValidateNull(rIS("ISvUnid")), "##,##0.0000")\r\n           vISuTrib = CStr(ValidateNull(rIS("uTrib_IS")))\r\n        End If',
    ),
    # F - Tb("uTrib_IS") assignment
    (
        'F - Tb("uTrib_IS") assignment',
        1,
        '    Tb("IS_vIS") = CDbl(Format(curISvIS, "0.00"))\r\nEnd Sub\r\n\r\nPrivate Sub Calcular_Total()',
        '    Tb("IS_vIS") = CDbl(Format(curISvIS, "0.00"))\r\n    Tb("uTrib_IS") = vISuTrib\r\nEnd Sub\r\n\r\nPrivate Sub Calcular_Total()',
    ),
]

for name, expected, old, new in patches:
    c = text.count(old)
    if c == expected:
        text = text.replace(old, new)
        print(f'  OK  {name}')
    else:
        print(f'ERRO  {name}: esperado {expected}, encontrado {c}')
        ok = False

if not ok:
    print('\nNFe_Completa.frm: ABORTADO — nenhuma alteracao salva')
    sys.exit(1)

out = text.encode('windows-1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(PATH_FRM, 'wb').write(out)
print('NFe_Completa.frm: salvo\n')

# ══════════════════════════════════════════════════════════════════
# modNFe.bas
# ══════════════════════════════════════════════════════════════════
PATH_BAS = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
text2 = open(PATH_BAS, 'rb').read().decode('windows-1252')

OLD_UTRIB = 'IIf(dISqUnid > 0, IIf(IsNull(NFeItens!UnidadeTributavel), "", NFeItens!UnidadeTributavel), "")'
NEW_UTRIB = 'IIf(IsNull(NFeItens!uTrib_IS), "", NFeItens!uTrib_IS)'

c = text2.count(OLD_UTRIB)
print(f'  G - uTrib_IS param: {c} ocorrencia(s) (esperado: 2)')
if c == 2:
    text2 = text2.replace(OLD_UTRIB, NEW_UTRIB)
    out2 = text2.encode('windows-1252')
    out2 = out2.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH_BAS, 'wb').write(out2)
    print('modNFe.bas: salvo\n')
else:
    print('modNFe.bas: ABORTADO')
    sys.exit(1)

print('=== patch_utrib_is: concluido ===')
