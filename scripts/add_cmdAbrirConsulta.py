# -*- coding: utf-8 -*-
"""
Adiciona o botao cmdAbrirConsulta na aba SITUACAO (Tab 0), logo apos
cmdExcluir, e registra no Tab(0).Control array (ControlCount 22 -> 23).
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

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


# 1) inserir o bloco de controle apos o End de cmdExcluir
i_begin = find_line_exact("      Begin ChamaleonBtn.chameleonButton cmdExcluir ")
i_end = find_line_exact("      End", i_begin)

novo_botao = """      Begin ChamaleonBtn.chameleonButton cmdAbrirConsulta
         Height          =   375
         Left            =   -63200
         TabIndex        =   235
         Top             =   5100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Consultar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End"""

lines[i_end] = "      End\r\n" + novo_botao

# 2) registrar no Tab(0).Control array
i_count = find_line_exact("      Tab(0).ControlCount=   22")
lines[i_count] = (
    '      Tab(0).Control(22)=   "cmdAbrirConsulta"\r\n'
    "      Tab(0).Control(22).Enabled=   0   'False\r\n"
    "      Tab(0).ControlCount=   23"
)

out_text = "\r\n".join(lines)
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - cmdAbrirConsulta adicionado")
