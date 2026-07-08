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

# ── 1: variáveis de módulo — acumuladores IBS ─────────────────────────────────
content = sub('Dim vAccumIBS',
    "Dim vCBSpRed As String           'bloco ibs_cbs",
    "Dim vCBSpRed As String           'bloco ibs_cbs" + R +
    "Dim vAccumIBS_BC  As Double      'arredondamento acumulado IBS" + R +
    "Dim vAccumIBS_UF  As Double      'arredondamento acumulado IBS" + R +
    "Dim vAccumIBS_Mun As Double      'arredondamento acumulado IBS",
    content)

# ── 2: cmdAdicionarItem_Click — ler acumulados ANTES do BeginTrans ────────────
content = sub('AccumIBS query antes BeginTrans',
    "sSQL = \"SELECT * FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text)" + R +
    "RsOpen Tb, sSQL" + R +
    R +
    "vgDb.BeginTrans",

    "sSQL = \"SELECT * FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text)" + R +
    "RsOpen Tb, sSQL" + R +
    R +
    "' Acumuladores IBS para arredondamento acumulado (itens ja gravados nesta nota)" + R +
    "Dim rAccumIBS As ADODB.Recordset" + R +
    "RsOpen rAccumIBS, \"SELECT ISNULL(SUM(IBS_vBC),0) AS sumBC, ISNULL(SUM(IBS_vIBSUF),0) AS sumUF, \" & _" + R +
    "                   \"ISNULL(SUM(IBS_vIBSMun),0) AS sumMun FROM NotaFiscalItens \" & _" + R +
    "                   \"WHERE CodigoNota = \" & Val(txtCodNota.Text)" + R +
    "vAccumIBS_BC  = CDbl(ValidateNull(rAccumIBS(\"sumBC\")))" + R +
    "vAccumIBS_UF  = CDbl(ValidateNull(rAccumIBS(\"sumUF\")))" + R +
    "vAccumIBS_Mun = CDbl(ValidateNull(rAccumIBS(\"sumMun\")))" + R +
    "If rAccumIBS.State <> 0 Then rAccumIBS.Close" + R +
    "Set rAccumIBS = Nothing" + R +
    R +
    "vgDb.BeginTrans",
    content)

# ── 3: Load_Data_Itens — substituir cálculo IBS por arredondamento acumulado ──
content = sub('IBS AccumRounding calc',
    "    curVIBSUF = CCur(Format(curBCCBSIBS * CDbl(IIf(vIBSUFpAliq = \"\", 0, vIBSUFpAliq)) / 100, \"0.00\"))" + R +
    "    curVIBSMun = CCur(Format(curBCCBSIBS * CDbl(IIf(vIBSMunpAliq = \"\", 0, vIBSMunpAliq)) / 100, \"0.00\"))",

    "    Dim dIBSUFRate As Double, dIBSMunRate As Double" + R +
    "    dIBSUFRate  = IIf(vIBSUFpAliq  = \"\", 0, Val(Replace(Replace(vIBSUFpAliq,  \".\", \"\"), \",\", \".\")))" + R +
    "    dIBSMunRate = IIf(vIBSMunpAliq = \"\", 0, Val(Replace(Replace(vIBSMunpAliq, \".\", \"\"), \",\", \".\")))" + R +
    "    Dim dNewTotalBC As Double" + R +
    "    dNewTotalBC = vAccumIBS_BC + CDbl(curBCCBSIBS)" + R +
    "    curVIBSUF  = CCur(CDbl(Format(dNewTotalBC * dIBSUFRate  / 100, \"0.00\")) - vAccumIBS_UF)" + R +
    "    curVIBSMun = CCur(CDbl(Format(dNewTotalBC * dIBSMunRate / 100, \"0.00\")) - vAccumIBS_Mun)",
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
