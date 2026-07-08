# -*- coding: utf-8 -*-
"""
Repara corrupcao de encoding em OS_Recapadora.frm: em algum momento desta
sessao, caracteres acentuados foram substituidos por U+FFFD codificado em
UTF-8 (bytes EF BF BD), que ao ser lido como cp1252 aparece como o texto
"ï¿½" (3 caracteres). O commit HEAD (antes desta sessao) nao tem nenhuma
ocorrencia disso - confirma que a corrupcao aconteceu durante a sessao.

Estrategia: alinhar as linhas do arquivo atual com as linhas do HEAD via
difflib (comparando "esqueletos" sem acentos, para casar linhas mesmo
corrompidas), e para blocos "equal"/"replace" de mesmo tamanho, sempre que
a linha atual tiver o padrao de corrupcao, substituir pela linha
correspondente do HEAD (que tem o acento correto).
"""
import difflib

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"
HEAD_PATH = r"C:\Users\NOTEBOOK\AppData\Local\Temp\claude\C--projeto\916fb1c0-4fd5-437b-8d03-a83de36ec5b2\scratchpad\orig_osrecap2.frm"

data_cur = open(PATH, "rb").read()
text_cur = data_cur.decode("cp1252")
lines_cur = text_cur.split("\r\n")

data_head = open(HEAD_PATH, "rb").read()
text_head = data_head.decode("cp1252")
lines_head = text_head.split("\n")

TARGET = chr(0xEF) + chr(0xBF) + chr(0xBD)


def skeleton(line):
    return "".join(c for c in line if ord(c) < 128)


sm = difflib.SequenceMatcher(
    a=[skeleton(l) for l in lines_head],
    b=[skeleton(l) for l in lines_cur],
    autojunk=False,
)
opcodes = sm.get_opcodes()

fixed = 0
unmatched_corrupted = []
new_lines = list(lines_cur)

for tag, i1, i2, j1, j2 in opcodes:
    if tag == "equal":
        for k in range(i2 - i1):
            hline = lines_head[i1 + k]
            cline = lines_cur[j1 + k]
            if TARGET in cline:
                new_lines[j1 + k] = hline
                fixed += 1
    elif tag == "replace":
        if (i2 - i1) == (j2 - j1):
            for k in range(i2 - i1):
                hline = lines_head[i1 + k]
                cline = lines_cur[j1 + k]
                if TARGET in cline:
                    new_lines[j1 + k] = hline
                    fixed += 1
        else:
            for k in range(j1, j2):
                if TARGET in lines_cur[k]:
                    unmatched_corrupted.append(k)
    elif tag == "insert":
        for k in range(j1, j2):
            if TARGET in lines_cur[k]:
                unmatched_corrupted.append(k)

print("fixed:", fixed)
print("unmatched corrupted lines:", len(unmatched_corrupted))
for k in unmatched_corrupted[:50]:
    print(" ", k + 1, repr(lines_cur[k][:150]))

out_text = "\r\n".join(new_lines)
out = out_text.encode("cp1252")

with open(PATH, "wb") as f:
    f.write(out)

print("saved, bytes:", len(out))
