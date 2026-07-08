data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# =============================================================================
# ALTERACAO 1: AplicarEstadoCheckboxes + chamadas em cboFinalidade_Click e Load_Controls
# =============================================================================

# 1a. Adicionar Dim bSupressChkEvents junto aos outros Dims de form-level
old_dim = b'Dim vTipoCRT As Integer\r\n'
new_dim = (
    b'Dim vTipoCRT As Integer\r\n'
    b'Dim bSupressChkEvents As Boolean\r\n'
)
c = data.count(old_dim)
print(f'1a. Dim bSupressChkEvents: {c}')
if c == 1: data = data.replace(old_dim, new_dim)

# 1b. Inserir sub AplicarEstadoCheckboxes antes de AplicarVisibilidadeGridItens
old_avsgi = b'Sub AplicarVisibilidadeGridItens()\r\n'
new_avsgi = (
    b'Sub AplicarEstadoCheckboxes()\r\n'
    b'    Dim bHabilitar As Boolean\r\n'
    b'    \' CRT 1/2/4 (Simples/MEI) com Finalidade <> 4: desabilita IPI, ST e RedBC\r\n'
    b'    bHabilitar = Not (vTipoCRT = 1 Or vTipoCRT = 2 Or vTipoCRT = 4) Or (Left(cboFinalidade.Text, 1) = "4")\r\n'
    b'    chkIPI.Enabled    = bHabilitar\r\n'
    b'    chkICMSST.Enabled = bHabilitar\r\n'
    b'    chkpRedBC.Enabled = bHabilitar\r\n'
    b'    If Not bHabilitar Then\r\n'
    b'        bSupressChkEvents = True\r\n'
    b'        chkIPI.Value    = 0\r\n'
    b'        chkICMSST.Value = 0\r\n'
    b'        chkpRedBC.Value = 0\r\n'
    b'        bSupressChkEvents = False\r\n'
    b'        AplicarVisibilidadeGridItens\r\n'
    b'    End If\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Sub AplicarVisibilidadeGridItens()\r\n'
)
c = data.count(old_avsgi)
print(f'1b. Inserir AplicarEstadoCheckboxes: {c}')
if c == 1: data = data.replace(old_avsgi, new_avsgi)

# 1c. cboFinalidade_Click: adicionar chamada a AplicarEstadoCheckboxes
old_fin = (
    b'Private Sub cboFinalidade_Click()\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    RecalcularItensNota\r\n'
    b'    CalcularICMSInterItensGERAL\r\n'
    b'End Sub\r\n'
)
new_fin = (
    b'Private Sub cboFinalidade_Click()\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    AplicarEstadoCheckboxes\r\n'
    b'    RecalcularItensNota\r\n'
    b'    CalcularICMSInterItensGERAL\r\n'
    b'End Sub\r\n'
)
c = data.count(old_fin)
print(f'1c. cboFinalidade_Click: {c}')
if c == 1: data = data.replace(old_fin, new_fin)

# 1d. Load_Controls: chamar AplicarEstadoCheckboxes antes de Exit Sub
old_lc = (
    b'    Mostrar_AliqUF\r\n'
    b'Exit Sub\r\n'
)
new_lc = (
    b'    Mostrar_AliqUF\r\n'
    b'    AplicarEstadoCheckboxes\r\n'
    b'Exit Sub\r\n'
)
c = data.count(old_lc)
print(f'1d. Load_Controls AplicarEstadoCheckboxes: {c}')
if c == 1: data = data.replace(old_lc, new_lc)

# =============================================================================
# ALTERACAO 2: chkIPI_Click, chkICMSST_Click, chkpRedBC_Click com recalculo
# =============================================================================
old_chk_ipi = b'Sub chkIPI_Click()\r\n   AplicarVisibilidadeGridItens\r\nEnd Sub\r\n'
new_chk_ipi = (
    b'Sub chkIPI_Click()\r\n'
    b'    If bSupressChkEvents Then Exit Sub\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    If chkIPI.Value = 0 And txtCodNota.Text <> "" Then\r\n'
    b'        dbData.Execute "UPDATE NotaFiscalItens SET IPIcEnq = \'999\', IPIvBC = 0, IPIpIPI = 0, IPIvIPI = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'    End If\r\n'
    b'    RecalcularItensNota\r\n'
    b'End Sub\r\n'
)
c = data.count(old_chk_ipi)
print(f'2a. chkIPI_Click: {c}')
if c == 1: data = data.replace(old_chk_ipi, new_chk_ipi)

old_chk_st = b'Sub chkICMSST_Click()\r\n   AplicarVisibilidadeGridItens\r\nEnd Sub\r\n'
new_chk_st = (
    b'Sub chkICMSST_Click()\r\n'
    b'    If bSupressChkEvents Then Exit Sub\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    RecalcularItensNota\r\n'
    b'End Sub\r\n'
)
c = data.count(old_chk_st)
print(f'2b. chkICMSST_Click: {c}')
if c == 1: data = data.replace(old_chk_st, new_chk_st)

old_chk_red = b'Sub chkpRedBC_Click()\r\n   AplicarVisibilidadeGridItens\r\nEnd Sub\r\n'
new_chk_red = (
    b'Sub chkpRedBC_Click()\r\n'
    b'    If bSupressChkEvents Then Exit Sub\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    If chkpRedBC.Value = 0 And txtCodNota.Text <> "" Then\r\n'
    b'        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'    End If\r\n'
    b'    RecalcularItensNota\r\n'
    b'End Sub\r\n'
)
c = data.count(old_chk_red)
print(f'2c. chkpRedBC_Click: {c}')
if c == 1: data = data.replace(old_chk_red, new_chk_red)

# =============================================================================
# ALTERACAO 3: Col 0 com numero da linha em FormatarGridItensNota
# =============================================================================
old_col0 = (
    b'      .rows = .rows - 1\r\n'
    b'\r\n'
    b"      'EAN em negrito\r\n"
)
new_col0 = (
    b'      .rows = .rows - 1\r\n'
    b'\r\n'
    b"      'Numero da linha no col 0\r\n"
    b'      For i = 1 To .rows - 1\r\n'
    b'         .TextMatrix(i, 0) = i\r\n'
    b'      Next i\r\n'
    b'\r\n'
    b"      'EAN em negrito\r\n"
)
c = data.count(old_col0)
print(f'3. Col 0 linha numero: {c}')
if c == 1: data = data.replace(old_col0, new_col0)

# =============================================================================
# ALTERACAO 4: Case 5 UND - validar contra lista aprovada
# =============================================================================
old_und = (
    b"    Case 5 ' UND\r\n"
    b'        If sVal = "" Then\r\n'
    b'            MsgBox "Unidade n\xe3o pode ser vazia!", vbInformation, "Aviso"\r\n'
    b'            Exit Sub\r\n'
    b'        End If\r\n'
    b'        If Len(sVal) > 2 Then\r\n'
    b'            MsgBox "Unidade deve ter no m\xe1ximo 2 caracteres!", vbInformation, "Aviso"\r\n'
    b'            Exit Sub\r\n'
    b'        End If\r\n'
    b'        sVal = UCase(sVal)\r\n'
)
new_und = (
    b"    Case 5 ' UND\r\n"
    b'        If sVal = "" Then\r\n'
    b'            MsgBox "Unidade n\xe3o pode ser vazia!", vbInformation, "Aviso"\r\n'
    b'            Exit Sub\r\n'
    b'        End If\r\n'
    b'        sVal = UCase(sVal)\r\n'
    b'        Dim sListaUND As String\r\n'
    b'        sListaUND = "|UN|PC|KG|CX|PA|PT|LT|ML|GR|DZ|FD|RL|JG|KT|LA|GL|BD|SC|PR|M2|M3|CT|EX|BJ|DI|MET|"\r\n'
    b'        If InStr(sListaUND, "|" & sVal & "|") = 0 Then\r\n'
    b'            MsgBox "Unidade \'" & sVal & "\' inv\xe1lida!" & vbCrLf & "Aceitas: UN PC KG CX PA PT LT ML GR DZ FD RL JG KT LA GL BD SC PR M2 M3 CT EX BJ DI MET", vbInformation, "Aviso"\r\n'
    b'            Exit Sub\r\n'
    b'        End If\r\n'
)
c = data.count(old_und)
print(f'4. Case 5 UND lista: {c}')
if c == 1: data = data.replace(old_und, new_und)

# Normalizar CRLF
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
