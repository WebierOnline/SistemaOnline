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

# ── 1: SELECT Mostrar_Aliquotas_Produto — adicionar cClassTrib ────────────────
content = sub('SELECT cClassTrib',
    ', IBSCBSCST, ISCST, cClassTrib_IS, fator_conversao_IS,',
    ', cClassTrib, IBSCBSCST, ISCST, cClassTrib_IS, fator_conversao_IS,',
    content)

# ── 2: Dim module-level vClassTrib (bloco ibs_cbs) ───────────────────────────
content = sub('Dim vClassTrib',
    "Dim vCBSpRed As String           'bloco ibs_cbs" + R +
    "Dim vAccumIBS_BC",
    "Dim vCBSpRed As String           'bloco ibs_cbs" + R +
    "Dim vClassTrib As String             'bloco ibs_cbs" + R +
    "Dim vAccumIBS_BC",
    content)

# ── 3: Dim module-level vClassTrib_IS (bloco IS) ─────────────────────────────
content = sub('Dim vClassTrib_IS',
    "Dim vISCST As String            'bloco IS" + R +
    "Dim vISTipoCalc",
    "Dim vISCST As String            'bloco IS" + R +
    "Dim vClassTrib_IS As String          'bloco IS" + R +
    "Dim vISTipoCalc",
    content)

# ── 4: Ler vClassTrib do recordset ───────────────────────────────────────────
content = sub('Read vClassTrib',
    "     vIBSCBSCST = ValidateNull(r(\"IBSCBSCST\"))" + R +
    "     " + R +
    "     ' CBS:",
    "     vIBSCBSCST = ValidateNull(r(\"IBSCBSCST\"))" + R +
    "     vClassTrib = Trim(ValidateNull(r(\"cClassTrib\")))" + R +
    "     " + R +
    "     ' CBS:",
    content)

# ── 5: Ler vClassTrib_IS do recordset ────────────────────────────────────────
content = sub('Read vClassTrib_IS',
    "     vISCST = ValidateNull(r(\"ISCST\"))" + R +
    "     vISTipoCalc = \"\"",
    "     vISCST = ValidateNull(r(\"ISCST\"))" + R +
    "     vClassTrib_IS = Trim(ValidateNull(r(\"cClassTrib_IS\")))" + R +
    "     vISTipoCalc = \"\"",
    content)

# ── 6: Reset vClassTrib/vClassTrib_IS — bloco sem indent (cboDescricao sub) ──
content = sub('Reset vClassTrib (no indent)',
    "vCBSpRed = \"\"" + R +
    "vISCST = \"\"" + R +
    "vISTipoCalc = \"\"",
    "vCBSpRed = \"\"" + R +
    "vClassTrib = \"\"" + R +
    "vISCST = \"\"" + R +
    "vClassTrib_IS = \"\"" + R +
    "vISTipoCalc = \"\"",
    content)

# ── 7: Reset vClassTrib/vClassTrib_IS — bloco com indent (else Mostrar_Aliq) ─
content = sub('Reset vClassTrib (indented)',
    "     vCBSpRed = \"\"" + R +
    "     vISCST = \"\"" + R +
    "     vISTipoCalc = \"\"",
    "     vCBSpRed = \"\"" + R +
    "     vClassTrib = \"\"" + R +
    "     vISCST = \"\"" + R +
    "     vClassTrib_IS = \"\"" + R +
    "     vISTipoCalc = \"\"",
    content)

# ── 8: Tb("cClassTrib") antes de Tb("IBSCBS_CST") ────────────────────────────
content = sub('Tb cClassTrib',
    "    Tb(\"IBSCBS_CST\") = Format(vIBSCBSCST, \"@\")",
    "    Tb(\"cClassTrib\") = Format(vClassTrib, \"@\")" + R +
    "    Tb(\"IBSCBS_CST\") = Format(vIBSCBSCST, \"@\")",
    content)

# ── 9: Tb("cClassTrib_IS") antes de Tb("IS_CST") ─────────────────────────────
content = sub('Tb cClassTrib_IS',
    "    Tb(\"IS_CST\") = Format(vISCST, \"@\")",
    "    Tb(\"cClassTrib_IS\") = Format(vClassTrib_IS, \"@\")" + R +
    "    Tb(\"IS_CST\") = Format(vISCST, \"@\")",
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
