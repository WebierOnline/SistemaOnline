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

# ── 1: SELECT produtos — trocar ISpIS por cClassTrib_IS ───────────────────────
content = sub('SELECT produtos ISpIS->cClassTrib_IS',
    'IBSCBSCST, ISCST, ISpIS,',
    'IBSCBSCST, ISCST, cClassTrib_IS,',
    content)

# ── 2: Bloco IS — substituir placeholders pelo lookup em tbISClassTrib ────────
old_is = (
    '     \' BLOCO IS: dados de duas tabelas (produtos )' + R +
    '     vISCST = ValidateNull(r("ISCST"))' + R +
    '     vISTipoCalc' + R +
    '     vISpAliq = Format(ValidateNull(r("ISpIS")), "##,##0.00")' + R +
    '     vISqUnid' + R +
    '     vISvUnid'
)

new_is = (
    '     \' IS: CST do produto; demais campos de tbISClassTrib (vigencia pelo periodo)' + R +
    '     vISCST = ValidateNull(r("ISCST"))' + R +
    '     vISTipoCalc = ""' + R +
    '     vISpAliq    = "0,00"' + R +
    '     vISqUnid    = "0,0000"' + R +
    '     vISvUnid    = "0,0000"' + R +
    '     Dim sClassIS As String' + R +
    '     sClassIS = Trim(ValidateNull(r("cClassTrib_IS")))' + R +
    '     If Len(sClassIS) > 0 Then' + R +
    '        Dim rIS As ADODB.Recordset' + R +
    '        RsOpen rIS, "SELECT TOP 1 tipo_calculo_is, ISpAliq, ISqUnid, ISvUnid " & _' + R +
    '                    "FROM tbISClassTrib " & _' + R +
    '                    "WHERE cClassTrib_IS = \'" & Replace(sClassIS, "\'", "\'\'") & "\' " & _' + R +
    '                    "ORDER BY CASE WHEN GETDATE() BETWEEN dIniVig AND dFimVig THEN 0 ELSE 1 END ASC, dFimVig DESC"' + R +
    '        If Not rIS.BOF Then' + R +
    '           vISTipoCalc = CStr(ValidateNull(rIS("tipo_calculo_is")))' + R +
    '           vISpAliq    = Format(ValidateNull(rIS("ISpAliq")), "##,##0.00")' + R +
    '           vISqUnid    = Format(ValidateNull(rIS("ISqUnid")), "##,##0.0000")' + R +
    '           vISvUnid    = Format(ValidateNull(rIS("ISvUnid")), "##,##0.0000")' + R +
    '        End If' + R +
    '        If rIS.State <> 0 Then rIS.Close' + R +
    '        Set rIS = Nothing' + R +
    '     End If'
)

content = sub('IS bloco tbISClassTrib', old_is, new_is, content)

# ── 3: Bloco Else — limpar todas as variaveis IS ──────────────────────────────
content = sub('Else clear IS vars',
    '     vISCST = ""' + R +
    '     vISpAliq = ""',
    '     vISCST      = ""' + R +
    '     vISTipoCalc = ""' + R +
    '     vISpAliq    = ""' + R +
    '     vISqUnid    = ""' + R +
    '     vISvUnid    = ""',
    content)

# ── 4: LimparObjetosProduto — limpar todas as variaveis IS ────────────────────
content = sub('LimparObjetos clear IS vars',
    'vISCST = ""' + R +
    'vISpAliq = ""' + R +
    'End Sub',
    'vISCST      = ""' + R +
    'vISTipoCalc = ""' + R +
    'vISpAliq    = ""' + R +
    'vISqUnid    = ""' + R +
    'vISvUnid    = ""' + R +
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
