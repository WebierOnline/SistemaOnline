# -*- coding: utf-8 -*-
# Patch: Produtos_Cadastro.frm
# - Adiciona BuscarDescricaoNCM que preenche lblNCMDescricao a partir de tbNCM
# - Chama em txtNCM_LostFocus e em MostrarDados_Produto

import os

FRM = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'

with open(FRM, 'rb') as f:
    data = f.read()

def norm(d):
    return d.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')

data = norm(data)

def patch(old_bytes, new_bytes, label):
    global data
    count = data.count(old_bytes)
    if count == 0:
        raise Exception(f'[{label}] Trecho nao encontrado')
    if count > 1:
        raise Exception(f'[{label}] Trecho nao e unico ({count} ocorrencias)')
    data = data.replace(old_bytes, new_bytes)
    print(f'OK: {label}')

enc = 'windows-1252'

# ── 1. txtNCM_LostFocus: no caminho invalido, limpar label ──────────────────
patch(
    ('        MsgBox "NCM Inv\xe1lido!", vbInformation, "Aviso do Sistema"\r\n'
     '        \'txtNCM.SetFocus\r\n'
     '        Exit Sub').encode(enc),
    ('        MsgBox "NCM Inv\xe1lido!", vbInformation, "Aviso do Sistema"\r\n'
     '        \'txtNCM.SetFocus\r\n'
     '        lblNCMDescricao.Caption = ""\r\n'
     '        Exit Sub').encode(enc),
    'LostFocus invalido: limpar lblNCMDescricao'
)

# ── 2. txtNCM_LostFocus: ao final (caminho valido e vazio), buscar descricao ─
patch(
    ('If txtNCM.Text <> "" Then\r\n'
     '    If Len(txtNCM.Text) < 8 Or Len(txtNCM.Text) > 8 Then\r\n'
     '        MsgBox "NCM Inv\xe1lido!", vbInformation, "Aviso do Sistema"\r\n'
     '        \'txtNCM.SetFocus\r\n'
     '        lblNCMDescricao.Caption = ""\r\n'
     '        Exit Sub\r\n'
     '    End If\r\n'
     'End If\r\n'
     'End Sub').encode(enc),
    ('If txtNCM.Text <> "" Then\r\n'
     '    If Len(txtNCM.Text) < 8 Or Len(txtNCM.Text) > 8 Then\r\n'
     '        MsgBox "NCM Inv\xe1lido!", vbInformation, "Aviso do Sistema"\r\n'
     '        \'txtNCM.SetFocus\r\n'
     '        lblNCMDescricao.Caption = ""\r\n'
     '        Exit Sub\r\n'
     '    End If\r\n'
     'End If\r\n'
     'BuscarDescricaoNCM\r\n'
     'End Sub').encode(enc),
    'LostFocus: chamar BuscarDescricaoNCM ao final'
)

# ── 3. MostrarDados_Produto: apos setar txtNCM, buscar descricao ────────────
patch(
    '    txtNCM.Text = ValidateNull(r("NCM"))\r\n'.encode(enc),
    ('    txtNCM.Text = ValidateNull(r("NCM"))\r\n'
     '    BuscarDescricaoNCM\r\n').encode(enc),
    'MostrarDados_Produto: chamar BuscarDescricaoNCM'
)

# ── 4. Adicionar sub BuscarDescricaoNCM antes de txtNCM_GotFocus ────────────
patch(
    'Private Sub txtNCM_GotFocus()'.encode(enc),
    ('Private Sub BuscarDescricaoNCM()\r\n'
     '    If Len(Trim(txtNCM.Text)) = 8 Then\r\n'
     '        lblNCMDescricao.Caption = SQLExecutaRetorno( _\r\n'
     '            "SELECT descricao FROM tbNCM WHERE NCM=\'" & Trim(txtNCM.Text) & "\'", _\r\n'
     '            "descricao", "")\r\n'
     '    Else\r\n'
     '        lblNCMDescricao.Caption = ""\r\n'
     '    End If\r\n'
     'End Sub\r\n'
     '\r\n'
     'Private Sub txtNCM_GotFocus()').encode(enc),
    'Adicionar sub BuscarDescricaoNCM'
)

data = norm(data)
with open(FRM, 'wb') as f:
    f.write(data)
print('Patch concluido:', FRM)
