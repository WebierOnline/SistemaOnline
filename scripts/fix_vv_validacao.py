path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: txtMargemVV_LostFocus -- bloqueia margem negativa ----
old1 = (
    'If txtMargemVV.Text = "" Then txtMargemVV.Text = 0\r\n'
    'varMargemVV = txtMargemVV.Text\r\n'
    '\r\n'
    'txtMargemVV.Text = FormatNumber(varMargemVV, 2) & "%"\r\n'
)
new1 = (
    'If txtMargemVV.Text = "" Then txtMargemVV.Text = 0\r\n'
    'varMargemVV = txtMargemVV.Text\r\n'
    '\r\n'
    'If varMargemVV < 0 Then\r\n'
    '   MsgBox "Margem negativa indica venda abaixo do custo." & vbCrLf & "A margem foi zerada.", vbExclamation, "Aviso"\r\n'
    '   varMargemVV = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemVV.Text = FormatNumber(varMargemVV, 2) & "%"\r\n'
)

# ---- R2: txtValorVV_LostFocus -- bloqueia valor < custo e evita div/0 ----
old2 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorVV.Text\r\n'
    'c = ((B - a) / a) * 100\r\n'
    '\r\n'
    'txtMargemVV.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorVV.Text = Format(txtValorVV.Text, ocMONEY)\r\n'
    'End Sub\r\n'
)
new2 = (
    'a = txtCusto.Text\r\n'
    'B = txtValorVV.Text\r\n'
    '\r\n'
    'If B < a Then\r\n'
    '   MsgBox "Valor de venda (" & Format(B, ocMONEY) & ") menor que o custo (" & Format(a, ocMONEY) & ")." & vbCrLf & "O valor foi corrigido para o custo.", vbExclamation, "Aviso"\r\n'
    '   B = a\r\n'
    '   txtValorVV.Text = Format(a, ocMONEY)\r\n'
    'End If\r\n'
    '\r\n'
    'If a > 0 Then\r\n'
    '   c = ((B - a) / a) * 100\r\n'
    'Else\r\n'
    '   c = 0\r\n'
    'End If\r\n'
    '\r\n'
    'txtMargemVV.Text = FormatNumber(c, 2) & "%"\r\n'
    'txtValorVV.Text = Format(B, ocMONEY)\r\n'
    'End Sub\r\n'
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))

data2 = data.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
