path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

old = (
    "Private Sub Menu_Ajuda_BaixarXML1_Click()\r\n"
    "Dim url As String\r\n"
    " ' O link do seu v\ufffdeo que voc\ufffd postou\r\n"
    " url = \"https://www.youtube.com/watch?v=A48N8kGpt9U\"\r\n"
    " \r\n"
    " \r\n"
    " ' Aqui usamos a constante 'conSwNormal' que voc\ufffd j\ufffd tem declarada\r\n"
    " ShellExecute Me.hwnd, \"open\", url, vbNullString, vbNullString, conSwNormal\r\n"
    "End Sub\r\n"
)

# Verifica com os caracteres exatos do arquivo
idx = data.find("Menu_Ajuda_BaixarXML1_Click()\r\n")
block = data[idx - len("Private Sub ") : idx - len("Private Sub ") + 400]
end = block.find("End Sub\r\n") + len("End Sub\r\n")
old_exact = block[:end]
print('old_exact:', repr(old_exact))
print('found by exact:', data.count(old_exact))

new = (
    "Private Sub Menu_Ajuda_BaixarXML1_Click()\r\n"
    "   Dim url As String\r\n"
    "   Dim chromePath As String\r\n"
    "   Dim firefoxPath As String\r\n"
    "\r\n"
    "   url = \"https://www.youtube.com/embed/A48N8kGpt9U?autoplay=1&rel=0&mute=1\"\r\n"
    "\r\n"
    "   ' Chrome: tenta Program Files depois Program Files (x86)\r\n"
    "   chromePath = \"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe\"\r\n"
    "   If Not Existe(chromePath) Then\r\n"
    "      chromePath = \"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe\"\r\n"
    "   End If\r\n"
    "\r\n"
    "   ' Firefox: tenta Program Files depois Program Files (x86)\r\n"
    "   firefoxPath = \"C:\\Program Files\\Mozilla Firefox\\firefox.exe\"\r\n"
    "   If Not Existe(firefoxPath) Then\r\n"
    "      firefoxPath = \"C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe\"\r\n"
    "   End If\r\n"
    "\r\n"
    "   If Existe(chromePath) Then\r\n"
    "      ShellExecute Me.hwnd, \"open\", chromePath, url, vbNullString, conSwNormal\r\n"
    "   ElseIf Existe(firefoxPath) Then\r\n"
    "      ShellExecute Me.hwnd, \"open\", firefoxPath, url, vbNullString, conSwNormal\r\n"
    "   Else\r\n"
    "      ShellExecute Me.hwnd, \"open\", url, vbNullString, vbNullString, conSwNormal\r\n"
    "   End If\r\n"
    "End Sub\r\n"
)

data2 = data.replace(old_exact, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
