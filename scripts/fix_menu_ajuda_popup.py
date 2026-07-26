path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

old = '   url = "https://www.youtube.com/watch?v=A48N8kGpt9U&autoplay=1&rel=0&mute=1"\r\n'
new = '   url = "https://www.youtube.com/watch_popup?v=A48N8kGpt9U&autoplay=1&rel=0&mute=1"\r\n'

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
