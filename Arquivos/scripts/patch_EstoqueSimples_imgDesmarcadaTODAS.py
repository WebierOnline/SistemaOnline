path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c, replace_all=False):
    n = c.count(old)
    if replace_all:
        if n == 0:
            errors.append(f'{label}: count=0')
            return c
        print(f'{label} OK ({n}x)')
        return c.replace(old, new)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# 1: sem indent — ImgMarcadaTODAS.Visible=False → adicionar imgDesmarcadaTODAS.Visible=True
#    (ImgMarcadaTODAS_Click e ResetarMarcas)
content = sub('Visible=False no-indent + desmarcadaTODAS',
    'ImgMarcadaTODAS.Visible = False\r\n'
    'lblMarcarTodas',
    'ImgMarcadaTODAS.Visible = False\r\n'
    'imgDesmarcadaTODAS.Visible = True\r\n'
    'lblMarcarTodas',
    content, replace_all=True)

# 2: 3-space indent — ImgMarcadaTODAS.Visible=False → adicionar imgDesmarcadaTODAS.Visible=True
#    (Formatar_Grid, Formatar_Grid_Fiscal, lblMarcarTodas_Click unmark branch)
content = sub('Visible=False 3sp + desmarcadaTODAS',
    '   ImgMarcadaTODAS.Visible = False\r\n'
    '   lblMarcarTodas',
    '   ImgMarcadaTODAS.Visible = False\r\n'
    '   imgDesmarcadaTODAS.Visible = True\r\n'
    '   lblMarcarTodas',
    content, replace_all=True)

# 3: 3-space indent — ImgMarcadaTODAS.Visible=True → adicionar imgDesmarcadaTODAS.Visible=False
#    (lblMarcarTodas_Click mark branch)
content = sub('Visible=True 3sp + desmarcadaTODAS=False',
    '   ImgMarcadaTODAS.Visible = True\r\n'
    '   lblMarcarTodas',
    '   ImgMarcadaTODAS.Visible = True\r\n'
    '   imgDesmarcadaTODAS.Visible = False\r\n'
    '   lblMarcarTodas',
    content)

# 4: Adicionar Private Sub imgDesmarcadaTODAS_Click antes de ImgMarcadaTODAS_Click
content = sub('add imgDesmarcadaTODAS_Click',
    'Private Sub ImgMarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'ImgMarcadaTODAS.Visible = False\r\n',

    'Private Sub imgDesmarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'imgDesmarcadaTODAS.Visible = False\r\n'
    'ImgMarcadaTODAS.Visible = True\r\n'
    'lblMarcarTodas.Caption = "Desmarcar todos"\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   Grid.TextMatrix(i, 0) = "1"\r\n'
    '   Grid.Row = i: Grid.Col = 0\r\n'
    '   Set Grid.CellPicture = ImgMarcada.Picture\r\n'
    '   Grid.CellPictureAlignment = 4\r\n'
    'Next i\r\n'
    'AvaliarFrmEdicao\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub ImgMarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'ImgMarcadaTODAS.Visible = False\r\n',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
