# -*- coding: utf-8 -*-
"""
Adiciona os criterios DATA, PERIODO e MENSAL em OS_Consulta.frm:
- novos controles (mskDataConsulta, mskPeriodoInicio/Fim, lblPeriodoAte,
  cboMesConsulta, cboAnoConsulta) dentro de Frame2, ocultos por padrao
- Preencher_Criterios com os 3 novos itens
- cboConsultaCriterios_Click/_Validate reescritos com um helper
  AtualizarCamposCriterios que mostra/esconde os campos certos
- cboMesConsulta_GotFocus / cboAnoConsulta_GotFocus (populam as combos)
- MostrarGrid_OS: 3 novos ramos ElseIf (DATA/PERIODO/MENSAL) em cada uma
  das 3 secoes de vTipoOS (Automoveis/Motos/Recapadora; Informatica/
  Celular; Comunicacao Visual), filtrando por os.DATA_ENTRADA
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


def find_block(prefix, start=0):
    i = None
    for j in range(start, len(lines)):
        if lines[j].lstrip().startswith(prefix):
            i = j
            break
    if i is None:
        raise SystemExit(f"ERRO: controle nao encontrado: {prefix!r}")
    indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
    end_marker = indent + "End"
    for k in range(i + 1, len(lines)):
        if lines[k] == end_marker:
            return i, k
    raise SystemExit(f"ERRO: End nao encontrado para {prefix!r}")


# ---------------------------------------------------------------
# 1) Object reference do msmask32.ocx (MaskEdBox)
# ---------------------------------------------------------------
i = find_line_exact('Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"')
lines[i] = (
    lines[i]
    + '\r\nObject = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"'
)

# ---------------------------------------------------------------
# 2) Novos controles apos o End de cboLocalizar
# ---------------------------------------------------------------
s, e = find_block("Begin VB.ComboBox cboLocalizar")
novos_controles = """         Begin MSMask.MaskEdBox mskDataConsulta
            Height          =   315
            Left            =   7860
            TabIndex        =   210
            Top             =   480
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskPeriodoInicio
            Height          =   315
            Left            =   7860
            TabIndex        =   211
            Top             =   480
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPeriodoAte
            AutoSize        =   -1  'True
            Caption         =   "at\xe9"
            Height          =   195
            Left            =   9195
            TabIndex        =   212
            Top             =   540
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskPeriodoFim
            Height          =   315
            Left            =   9540
            TabIndex        =   213
            Top             =   480
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboMesConsulta
            Height          =   315
            Left            =   7860
            TabIndex        =   214
            Top             =   480
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.ComboBox cboAnoConsulta
            Height          =   315
            Left            =   9600
            Sorted          =   -1  'True
            TabIndex        =   215
            Top             =   480
            Visible         =   0   'False
            Width           =   1000
         End""".replace("\xe9", "é")
lines[e] = lines[e] + "\r\n" + novos_controles

# ---------------------------------------------------------------
# 3) Preencher_Criterios
# ---------------------------------------------------------------
i = find_line_exact("Private Sub Preencher_Criterios()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
expected = [
    "Private Sub Preencher_Criterios()",
    "cboConsultaCriterios.Clear",
    'cboConsultaCriterios.AddItem "TODOS"',
    'cboConsultaCriterios.AddItem "C\xd3D. OS"'.replace("\xd3", "Ó"),
    'cboConsultaCriterios.AddItem "CLIENTE"',
    "End Sub",
]
assert old == expected, old
novo = [
    "Private Sub Preencher_Criterios()",
    "cboConsultaCriterios.Clear",
    'cboConsultaCriterios.AddItem "TODOS"',
    'cboConsultaCriterios.AddItem "CÓD. OS"',
    'cboConsultaCriterios.AddItem "CLIENTE"',
    'cboConsultaCriterios.AddItem "DATA"',
    'cboConsultaCriterios.AddItem "PERÍODO"',
    'cboConsultaCriterios.AddItem "MENSAL"',
    "End Sub",
]
lines[i : end + 1] = novo

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - parte 1 (controles + Preencher_Criterios) aplicada")
