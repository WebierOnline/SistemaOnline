path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: declaracoes de modulo ──────────────────────────────────────────────────
content = sub('Dim vIBSpRed vCBSpRed',
    'Dim vIBSMunpAliq As String       \'bloco ibs_cbs' + R +
    'Dim vISCST As String            \'bloco IS',

    'Dim vIBSMunpAliq As String       \'bloco ibs_cbs' + R +
    'Dim vIBSpRed As String           \'bloco ibs_cbs' + R +
    'Dim vCBSpRed As String           \'bloco ibs_cbs' + R +
    'Dim vISCST As String            \'bloco IS',
    content)

# ── 2: LimparVariaveisItens — reset ──────────────────────────────────────────
content = sub('LimparVariaveisItens reset pRed',
    'vIBSMunpAliq = ""' + R +
    'vISCST = ""',

    'vIBSMunpAliq = ""' + R +
    'vIBSpRed = ""' + R +
    'vCBSpRed = ""' + R +
    'vISCST = ""',
    content)

# ── 3: Mostrar_Aliquotas_Produto — lookup TbIBSCBSClassTrib ──────────────────
content = sub('Mostrar_Aliquotas lookup IBSpRed',
    '           If rCid.State <> 0 Then rCid.Close' + R +
    '           Set rCid = Nothing' + R +
    '        End If' + R +
    '     End If' + R +
    '     ' + R +
    '     \' IS: CST do produto; demais campos de tbISClassTrib (vigencia pelo periodo)',

    '           If rCid.State <> 0 Then rCid.Close' + R +
    '           Set rCid = Nothing' + R +
    '        End If' + R +
    '     End If' + R +
    '     ' + R +
    '     \' IBS/CBS reducao: TbIBSCBSClassTrib (vigencia pelo periodo)' + R +
    '     vIBSpRed = "0,00"' + R +
    '     vCBSpRed = "0,00"' + R +
    '     Dim rIBSCBS As ADODB.Recordset' + R +
    '     RsOpen rIBSCBS, "SELECT TOP 1 pRedIBS, pRedCBS FROM TbIBSCBSClassTrib " & _' + R +
    '                     "INNER JOIN produtos ON produtos.cClassTrib = TbIBSCBSClassTrib.cClassTrib " & _' + R +
    '                     "WHERE produtos.codigo = " & Val(txtCodProduto.Text) & " " & _' + R +
    '                     "ORDER BY CASE WHEN GETDATE() BETWEEN dIniVig AND dFimVig THEN 0 ELSE 1 END ASC, dFimVig DESC"' + R +
    '     If Not rIBSCBS.BOF Then' + R +
    '        vIBSpRed = Format(ValidateNull(rIBSCBS("pRedIBS")), "##,##0.00")' + R +
    '        vCBSpRed = Format(ValidateNull(rIBSCBS("pRedCBS")), "##,##0.00")' + R +
    '     End If' + R +
    '     If rIBSCBS.State <> 0 Then rIBSCBS.Close' + R +
    '     Set rIBSCBS = Nothing' + R +
    '     ' + R +
    '     \' IS: CST do produto; demais campos de tbISClassTrib (vigencia pelo periodo)',
    content)

# ── 4: Mostrar_Aliquotas_Produto — else (produto nao encontrado) ──────────────
content = sub('Mostrar_Aliquotas else reset pRed',
    '     vIBSMunpAliq = ""' + R +
    '     vISCST = ""',

    '     vIBSMunpAliq = ""' + R +
    '     vIBSpRed = ""' + R +
    '     vCBSpRed = ""' + R +
    '     vISCST = ""',
    content)

# ── 5: Load_Data_Itens — usar vIBSpRed/vCBSpRed em vez de 0 ──────────────────
content = sub('Load_Data_Itens IBS_pRed real value',
    '    Tb("IBS_pRed") = CDbl(Format(0, "0.00"))',
    '    Tb("IBS_pRed") = CDbl(IIf(vIBSpRed = "", 0, Format(vIBSpRed, "@")))',
    content)

content = sub('Load_Data_Itens CBS_pRed real value',
    '    Tb("CBS_pRed") = CDbl(Format(0, "0.00"))',
    '    Tb("CBS_pRed") = CDbl(IIf(vCBSpRed = "", 0, Format(vCBSpRed, "@")))',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
