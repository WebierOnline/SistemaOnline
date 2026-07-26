# -*- coding: utf-8 -*-
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

DEST = 'O cadastro do DESTINAT\xc1RIO possui erros'

# ═══════════════════════════════════════════════════════════════════
# VerificarDestinatarioEnviar
# ═══════════════════════════════════════════════════════════════════

# 1: FORNECEDOR query — adicionar razao as vNome
content = sub('ForneQuery_Enviar_vNome',
    "    sSQL = \"SELECT *, 'JUR\xcdDICA' as vTipo  FROM FORNECEDOR WHERE codigo = \" & Val(vCodCliente)",
    "    sSQL = \"SELECT *, 'JUR\xcdDICA' as vTipo, razao as vNome FROM FORNECEDOR WHERE codigo = \" & Val(vCodCliente)",
    content)

# 2: cliente query — adicionar nome as vNome (contexto com blank-line + Set r sem blank)
content = sub('ClienteQuery_Enviar_vNome',
    "    sSQL = \"SELECT *, tipo as vTipo FROM cliente WHERE codigo = \" & Val(vCodCliente)" + R +
    "End If" + R + R +
    "Set r = dbData.OpenRecordset(sSQL)",
    "    sSQL = \"SELECT *, tipo as vTipo, nome as vNome FROM cliente WHERE codigo = \" & Val(vCodCliente)" + R +
    "End If" + R + R +
    "Set r = dbData.OpenRecordset(sSQL)",
    content)

# 3: endereco — IsEmpty → Vazio  (distinguido por "erros no [Campo:")
content = sub('IsEmpty_endereco_Enviar',
    'If IsEmpty(r("endereco")) Then If ShowMsg("' + DEST + ' no [Campo: Endere\xe7o]!',
    'If Vazio(r("endereco")) Then If ShowMsg("' + DEST + ' no [Campo: Endere\xe7o]!',
    content)

# 4: numero
content = sub('IsEmpty_numero_Enviar',
    'If IsEmpty(r("numero")) Then If ShowMsg("' + DEST + ' no [Campo: N\xfamero]!',
    'If Vazio(r("numero")) Then If ShowMsg("' + DEST + ' no [Campo: N\xfamero]!',
    content)

# 5: bairro — IsEmpty → Vazio + Len seguro para campos DAO nulos
content = sub('IsEmpty_bairro_Enviar',
    'If IsEmpty(r("bairro")) Or Len(r("bairro")) < 4 Then If ShowMsg("' + DEST + ' no [Campo: Bairro]!',
    'If Vazio(r("bairro")) Or Len(IIf(IsNull(r("bairro")), "", r("bairro"))) < 4 Then If ShowMsg("' + DEST + ' no [Campo: Bairro]!',
    content)

# 6: cidade
content = sub('IsEmpty_cidade_Enviar',
    'If IsEmpty(r("cidade")) Then If ShowMsg("' + DEST + ' no [Campo: Cidade]!',
    'If Vazio(r("cidade")) Then If ShowMsg("' + DEST + ' no [Campo: Cidade]!',
    content)

# 7: estado
content = sub('IsEmpty_estado_Enviar',
    'If IsEmpty(r("estado")) Then If ShowMsg("' + DEST + ' no [Campo: Estado]!',
    'If Vazio(r("estado")) Then If ShowMsg("' + DEST + ' no [Campo: Estado]!',
    content)

# 8: CodigoIBGE — remover "= 0" redundante, corrigir Len para DAO
content = sub('CodigoIBGE_Enviar',
    'If IsEmpty(r("CodigoIBGE")) Or r("CodigoIBGE") = "0" Or Len(r("CodigoIBGE")) <> 7 Then If ShowMsg("' + DEST + ' no [Campo: C\xf3d IBGE]!',
    'If Vazio(r("CodigoIBGE")) Or Len(CStr(IIf(IsNull(r("CodigoIBGE")), 0, r("CodigoIBGE")))) <> 7 Then If ShowMsg("' + DEST + ' no [Campo: C\xf3d IBGE]!',
    content)

# 9: CEP — < 10 → RemoverFormato <> 8
content = sub('CEP_Enviar',
    'If IsEmpty(r("CEP")) Or Len(r("CEP")) < 10 Then If ShowMsg("' + DEST + ' no [Campo: CEP]!',
    'If Vazio(r("CEP")) Or Len(RemoverFormato(IIf(IsNull(r("CEP")), "", CStr(r("CEP"))))) <> 8 Then If ShowMsg("' + DEST + ' no [Campo: CEP]!',
    content)

# 10: CPF/CNPJ block — IsEmpty(vCPF) → vCPF = ""
# (distinguido por: tem RURAL + vPossuiErro, ao contrário de VerificarDestinatario)
content = sub('CPF_CNPJ_block_Enviar',
    '    If r("TipoContribuinte") = 9 Then' + R +
    '        If IsEmpty(vCPF) Or Len(vCPF) < 11 Then If ShowMsg("' + DEST + ' no [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '    Else' + R +
    '        If r("vTipo") = "RURAL" Then' + R +
    '            If IsEmpty(vCPF) Or Len(vCPF) < 11 Then If ShowMsg("' + DEST + ' no [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '        Else' + R +
    '            If IsEmpty(vCPF) Or Len(vCPF) < 14 Then If ShowMsg("' + DEST + ' no [Campo: CNPJ]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '        End If' + R +
    '    End If',

    '    If r("TipoContribuinte") = 9 Then' + R +
    '        If vCPF = "" Or Len(vCPF) < 11 Then If ShowMsg("' + DEST + ' no [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '    Else' + R +
    '        If r("vTipo") = "RURAL" Then' + R +
    '            If vCPF = "" Or Len(vCPF) < 11 Then If ShowMsg("' + DEST + ' no [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '        Else' + R +
    '            If vCPF = "" Or Len(vCPF) < 14 Then If ShowMsg("' + DEST + ' no [Campo: CNPJ]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '        End If' + R +
    '    End If',
    content)

# 11: Inserir verificação de Nome/Razão Social antes do bloco TipoContribuinte=1
# Âncora: "no [Campo: Insc." distingue de VerificarDestinatario (" [Campo: Insc.")
content = sub('Add_nome_check_Enviar',
    '    ' + R +
    '    If r("TipoContribuinte") = 1 Then' + R +
    '        If Vazio(r("ie")) Then If ShowMsg("' + DEST + ' no [Campo: Insc. Estadual]!',
    '    ' + R +
    '    If Vazio(r("vNome")) Then If ShowMsg("' + DEST + ' no [Campo: Nome/Raz\xe3o Social]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then vPossuiErro = True: Exit Sub Else: vPossuiErro = True: GoTo AtualizarCliente' + R +
    '    If r("TipoContribuinte") = 1 Then' + R +
    '        If Vazio(r("ie")) Then If ShowMsg("' + DEST + ' no [Campo: Insc. Estadual]!',
    content)

# ═══════════════════════════════════════════════════════════════════
# VerificarDestinatario  (sem "no" antes de [Campo:, sem vPossuiErro)
# ═══════════════════════════════════════════════════════════════════

# 12: FORNECEDOR query — adicionar razao as vNome
content = sub('ForneQuery_Dest_vNome',
    "    sSQL = \"SELECT * FROM FORNECEDOR WHERE codigo = \" & Val(vCodCliente)" + R + "Else",
    "    sSQL = \"SELECT *, razao as vNome FROM FORNECEDOR WHERE codigo = \" & Val(vCodCliente)" + R + "Else",
    content)

# 13: cliente query — adicionar nome as vNome (sem blank-line antes do Set r)
content = sub('ClienteQuery_Dest_vNome',
    "    sSQL = \"SELECT * FROM cliente WHERE codigo = \" & Val(vCodCliente)" + R +
    "End If" + R +
    "Set r = dbData.OpenRecordset(sSQL)",
    "    sSQL = \"SELECT *, nome as vNome FROM cliente WHERE codigo = \" & Val(vCodCliente)" + R +
    "End If" + R +
    "Set r = dbData.OpenRecordset(sSQL)",
    content)

# 14: endereco (distinguido por "erros [Campo:" sem "no")
content = sub('IsEmpty_endereco_Dest',
    'If IsEmpty(r("endereco")) Then If ShowMsg("' + DEST + ' [Campo: Endere\xe7o]!',
    'If Vazio(r("endereco")) Then If ShowMsg("' + DEST + ' [Campo: Endere\xe7o]!',
    content)

# 15: numero
content = sub('IsEmpty_numero_Dest',
    'If IsEmpty(r("numero")) Then If ShowMsg("' + DEST + ' [Campo: N\xfamero]!',
    'If Vazio(r("numero")) Then If ShowMsg("' + DEST + ' [Campo: N\xfamero]!',
    content)

# 16: bairro (sem Len < 4 nesta rotina)
content = sub('IsEmpty_bairro_Dest',
    'If IsEmpty(r("bairro")) Then If ShowMsg("' + DEST + ' [Campo: Bairro]!',
    'If Vazio(r("bairro")) Then If ShowMsg("' + DEST + ' [Campo: Bairro]!',
    content)

# 17: cidade
content = sub('IsEmpty_cidade_Dest',
    'If IsEmpty(r("cidade")) Then If ShowMsg("' + DEST + ' [Campo: Cidade]!',
    'If Vazio(r("cidade")) Then If ShowMsg("' + DEST + ' [Campo: Cidade]!',
    content)

# 18: estado
content = sub('IsEmpty_estado_Dest',
    'If IsEmpty(r("estado")) Then If ShowMsg("' + DEST + ' [Campo: Estado]!',
    'If Vazio(r("estado")) Then If ShowMsg("' + DEST + ' [Campo: Estado]!',
    content)

# 19: CodigoIBGE — < 7 → <> 7, IsEmpty → Vazio, Len seguro
content = sub('CodigoIBGE_Dest',
    'If IsEmpty(r("CodigoIBGE")) Or Len(r("CodigoIBGE")) < 7 Then If ShowMsg("' + DEST + ' [Campo: C\xf3d IBGE]!',
    'If Vazio(r("CodigoIBGE")) Or Len(CStr(IIf(IsNull(r("CodigoIBGE")), 0, r("CodigoIBGE")))) <> 7 Then If ShowMsg("' + DEST + ' [Campo: C\xf3d IBGE]!',
    content)

# 20: CEP
content = sub('CEP_Dest',
    'If IsEmpty(r("CEP")) Or Len(r("CEP")) < 10 Then If ShowMsg("' + DEST + ' [Campo: CEP]!',
    'If Vazio(r("CEP")) Or Len(RemoverFormato(IIf(IsNull(r("CEP")), "", CStr(r("CEP"))))) <> 8 Then If ShowMsg("' + DEST + ' [Campo: CEP]!',
    content)

# 21: CPF/CNPJ block — IsEmpty(vCPF) → vCPF = ""  (sem RURAL, sem vPossuiErro)
content = sub('CPF_CNPJ_block_Dest',
    '    If r("TipoContribuinte") = 9 Then' + R +
    '        If IsEmpty(vCPF) Or Len(vCPF) < 11 Then If ShowMsg("' + DEST + ' [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente' + R +
    '    Else' + R +
    '        If IsEmpty(vCPF) Or Len(vCPF) < 14 Then If ShowMsg("' + DEST + ' [Campo: CNPJ]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente' + R +
    '    End If',

    '    If r("TipoContribuinte") = 9 Then' + R +
    '        If vCPF = "" Or Len(vCPF) < 11 Then If ShowMsg("' + DEST + ' [Campo: CPF]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente' + R +
    '    Else' + R +
    '        If vCPF = "" Or Len(vCPF) < 14 Then If ShowMsg("' + DEST + ' [Campo: CNPJ]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente' + R +
    '    End If',
    content)

# 22: Inserir verificação de Nome/Razão Social antes do bloco TipoContribuinte=1
# Âncora: " [Campo: Insc." sem "no" distingue de VerificarDestinatarioEnviar
content = sub('Add_nome_check_Dest',
    '    ' + R +
    '    If r("TipoContribuinte") = 1 Then' + R +
    '        If Vazio(r("ie")) Then If ShowMsg("' + DEST + ' [Campo: Insc. Estadual]!',
    '    ' + R +
    '    If Vazio(r("vNome")) Then If ShowMsg("' + DEST + ' [Campo: Nome/Raz\xe3o Social]!" & vbNewLine & "Deseja atualizar o cadastro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub Else: GoTo AtualizarCliente' + R +
    '    If r("TipoContribuinte") = 1 Then' + R +
    '        If Vazio(r("ie")) Then If ShowMsg("' + DEST + ' [Campo: Insc. Estadual]!',
    content)

# ═══════════════════════════════════════════════════════════════════
# cmdAdicionarItem_Click — exigir cliente selecionado
# ═══════════════════════════════════════════════════════════════════

# 23: bloco de guarda antes de adicionar item
content = sub('TxtCodCliente_check_AdicionarItem',
    'If txtSubTotal.Text = "" Then Exit Sub' + R +
    'If Len(vNCM) < 8 Then',
    'If txtSubTotal.Text = "" Then Exit Sub' + R +
    'If TxtCodCliente.Text = "" Then MsgBox "Selecione um cliente/fornecedor antes de adicionar itens.", vbExclamation, "Online Commerce": Exit Sub' + R +
    'VerificarDestinatarioEnviar' + R +
    'If vPossuiErro Then Exit Sub' + R +
    'If Len(vNCM) < 8 Then',
    content)

# ═══════════════════════════════════════════════════════════════════
if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
