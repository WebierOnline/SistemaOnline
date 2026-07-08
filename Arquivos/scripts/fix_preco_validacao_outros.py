path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ================================================================
# Margem: insere checagem de negativo antes do FormatNumber
# Valor:  insere checagem valor < custo e protecao div/0
# ================================================================

# ---- R1: txtMargemVP_LostFocus ----
old1 = (
    'txtMargemVP.Text = FormatNumber(varMargemVP, 2) & "%"\r\n'
    '\r\n'
    'CalcularPrecos\r\n'
    'lblAviso.Visible = False\r\n'
    'End Sub\r\n'
)
new1 = (
    'If varMargemVP < 0 Then\r\n'
    '   MsgBox "Margem negativa indica venda abaixo do custo." & vbCrLf & "A margem foi zerada.", vbExclamation, "Aviso"\r\n'
    '   varMargemVP = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemVP.Text = FormatNumber(varMargemVP, 2) & "%"\r\n'
    '\r\n'
    'CalcularPrecos\r\n'
    'lblAviso.Visible = False\r\n'
    'End Sub\r\n'
)

# ---- R2: txtMargemAV_LostFocus ----
old2 = (
    'txtMargemAV.Text = FormatNumber(varMargemAV, 2) & "%"\r\n'
    '\r\n'
    'CalcularPrecos\r\n'
    'lblAviso.Visible = False\r\n'
    'End Sub\r\n'
)
new2 = (
    'If varMargemAV < 0 Then\r\n'
    '   MsgBox "Margem negativa indica venda abaixo do custo." & vbCrLf & "A margem foi zerada.", vbExclamation, "Aviso"\r\n'
    '   varMargemAV = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemAV.Text = FormatNumber(varMargemAV, 2) & "%"\r\n'
    '\r\n'
    'CalcularPrecos\r\n'
    'lblAviso.Visible = False\r\n'
    'End Sub\r\n'
)

# ---- R3: txtMargemAP_LostFocus ----
old3 = (
    'txtMargemAP.Text = FormatNumber(varMargemAP, 2) & "%"\r\n'
    '\r\n'
    'CalcularPrecos\r\n'
    'lblAviso.Visible = False\r\n'
    'End Sub\r\n'
)
new3 = (
    'If varMargemAP < 0 Then\r\n'
    '   MsgBox "Margem negativa indica venda abaixo do custo." & vbCrLf & "A margem foi zerada.", vbExclamation, "Aviso"\r\n'
    '   varMargemAP = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemAP.Text = FormatNumber(varMargemAP, 2) & "%"\r\n'
    '\r\n'
    'CalcularPrecos\r\n'
    'lblAviso.Visible = False\r\n'
    'End Sub\r\n'
)

# ---- R4: txtValorVP_LostFocus ----
old4 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorVP.Text\r\n'
    'c = ((B - a) / a) * 100\r\n'
    '\r\n'
    'txtMargemVP.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorVP.Text = Format(txtValorVP.Text, ocMONEY)\r\n'
    'End Sub\r\n'
)
new4 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorVP.Text\r\n'
    '\r\n'
    'If B < a Then\r\n'
    '   MsgBox "Valor de venda (" & Format(B, ocMONEY) & ") menor que o custo (" & Format(a, ocMONEY) & ")." & vbCrLf & "O valor foi corrigido para o custo.", vbExclamation, "Aviso"\r\n'
    '   B = a\r\n'
    '   txtValorVP.Text = Format(a, ocMONEY)\r\n'
    'End If\r\n'
    '\r\n'
    'If a > 0 Then\r\n'
    '   c = ((B - a) / a) * 100\r\n'
    'Else\r\n'
    '   c = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemVP.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorVP.Text = Format(B, ocMONEY)\r\n'
    'End Sub\r\n'
)

# ---- R5: txtValorAV_LostFocus ----
old5 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorAV.Text\r\n'
    'c = ((B - a) / a) * 100\r\n'
    '\r\n'
    'txtMargemAV.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorAV.Text = Format(txtValorAV.Text, ocMONEY)\r\n'
    'End Sub\r\n'
)
new5 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorAV.Text\r\n'
    '\r\n'
    'If B < a Then\r\n'
    '   MsgBox "Valor de venda (" & Format(B, ocMONEY) & ") menor que o custo (" & Format(a, ocMONEY) & ")." & vbCrLf & "O valor foi corrigido para o custo.", vbExclamation, "Aviso"\r\n'
    '   B = a\r\n'
    '   txtValorAV.Text = Format(a, ocMONEY)\r\n'
    'End If\r\n'
    '\r\n'
    'If a > 0 Then\r\n'
    '   c = ((B - a) / a) * 100\r\n'
    'Else\r\n'
    '   c = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemAV.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorAV.Text = Format(B, ocMONEY)\r\n'
    'End Sub\r\n'
)

# ---- R6: txtValorAP_LostFocus ----
old6 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorAP.Text\r\n'
    'c = ((B - a) / a) * 100\r\n'
    '\r\n'
    'txtMargemAP.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorAP.Text = Format(txtValorAP.Text, ocMONEY)\r\n'
    'End Sub\r\n'
)
new6 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorAP.Text\r\n'
    '\r\n'
    'If B < a Then\r\n'
    '   MsgBox "Valor de venda (" & Format(B, ocMONEY) & ") menor que o custo (" & Format(a, ocMONEY) & ")." & vbCrLf & "O valor foi corrigido para o custo.", vbExclamation, "Aviso"\r\n'
    '   B = a\r\n'
    '   txtValorAP.Text = Format(a, ocMONEY)\r\n'
    'End If\r\n'
    '\r\n'
    'If a > 0 Then\r\n'
    '   c = ((B - a) / a) * 100\r\n'
    'Else\r\n'
    '   c = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemAP.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorAP.Text = Format(B, ocMONEY)\r\n'
    'End Sub\r\n'
)

replacements = [
    ('r1 txtMargemVP', old1, new1),
    ('r2 txtMargemAV', old2, new2),
    ('r3 txtMargemAP', old3, new3),
    ('r4 txtValorVP',  old4, new4),
    ('r5 txtValorAV',  old5, new5),
    ('r6 txtValorAP',  old6, new6),
]

data2 = data
for label, old, new in replacements:
    count = data2.count(old)
    print(f'{label} found: {count}')
    data2 = data2.replace(old, new, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
