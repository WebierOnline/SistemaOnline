# -*- coding: utf-8 -*-
"""
UX da atualizacao da tabela IBPT pela nuvem (menu manual):
- frmImportarIBPT.ImportarIBPTdeArquivo ganha 3o parametro opcional oProgresso (late-bound)
  pra reportar andamento pra uma forminha externa; grava mensagemErro no erro.
- Tela_Principal.Menu_FISCAL_Reforma_IBPTNuvem_Click: usa frmIBPTProgresso (barra) durante
  download E importacao; frmImportarIBPT nunca aparece; tudo fecha no fim + msg final.
.frm cp1252 -> edicao binaria + normalizacao CRLF.
"""
import sys

def norm(b): return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

def patch(path, repls):
    d = open(path, "rb").read().replace(b"\r\n", b"\n")
    for i, (old, new) in enumerate(repls, 1):
        if new in d and old not in d:
            print(f"  [{path}] #{i} ja aplicado"); continue
        if d.count(old) != 1:
            print(f"  [{path}] #{i}: {d.count(old)} ocorrencias (esperado 1)")
            sys.exit(1)
        d = d.replace(old, new, 1)
    open(path, "wb").write(norm(d))
    print(f"  OK {path}")


imp = r"C:\projeto\OnlineCommerce\Forms\frmImportarIBPT.frm"
patch(imp, [
    # assinatura: + oProgresso
    (
b'Public Function ImportarIBPTdeArquivo(ByVal caminhoCSV As String, Optional ByVal bSilencioso As Boolean = False) As Boolean',
b'Public Function ImportarIBPTdeArquivo(ByVal caminhoCSV As String, Optional ByVal bSilencioso As Boolean = False, Optional ByVal oProgresso As Object) As Boolean',
    ),
    # limpando tabela
    (
b'    If Not bSilencioso Then lblProgresso.Caption = "Limpando tabela...": DoEvents\n    dbData.Execute "DELETE FROM TabelaIBPT"',
b'    If Not bSilencioso Then lblProgresso.Caption = "Limpando tabela...": DoEvents\n    If Not (oProgresso Is Nothing) Then oProgresso.DefinirMensagem "Limpando a tabela atual..."\n    dbData.Execute "DELETE FROM TabelaIBPT"',
    ),
    # dentro do loop: reporta % pro oProgresso
    (
b'''        If Not bSilencioso Then
            If nLinha Mod 100 = 0 Then
                lblProgresso.Caption = "Importando: " & nLinha & " de " & nTotal & " registros..."
                DoEvents
            End If
        End If
ProxLinhaAuto:''',
b'''        If Not bSilencioso Then
            If nLinha Mod 100 = 0 Then
                lblProgresso.Caption = "Importando: " & nLinha & " de " & nTotal & " registros..."
                DoEvents
            End If
        End If
        If Not (oProgresso Is Nothing) Then
            If nLinha Mod 100 = 0 Then oProgresso.DefinirProgresso nLinha, nTotal
        End If
ProxLinhaAuto:''',
    ),
    # sincronizando tbNCM
    (
b'    If Not bSilencioso Then lblProgresso.Caption = "Sincronizando tbNCM...": DoEvents',
b'    If Not bSilencioso Then lblProgresso.Caption = "Sincronizando tbNCM...": DoEvents\n    If Not (oProgresso Is Nothing) Then oProgresso.DefinirMensagem "Sincronizando a tabela NCM..."',
    ),
    # erro: grava mensagemErro pra quem chamou (inclusive modo silencioso) poder exibir
    (
b'''ErrAuto:
    Dim sErrA As String
    sErrA = Err.Description
    On Error Resume Next
    Close #iFile''',
b'''ErrAuto:
    Dim sErrA As String
    sErrA = Err.Description
    mensagemErro = sErrA
    On Error Resume Next
    Close #iFile''',
    ),
])


tp = r"C:\projeto\OnlineCommerce\Forms\Tela_Principal.frm"
patch(tp, [
    (
b'''Private Sub Menu_FISCAL_Reforma_IBPTNuvem_Click()
   ' baixa a tabela IBPT mais recente do Shared Drive "IBPT" e importa na hora (sem gate de data)
   If MsgBox("Baixar a tabela IBPT mais recente da nuvem e atualizar agora?", vbQuestion + vbYesNo, "Tabela IBPT") <> vbYes Then Exit Sub

   Dim sPastaDL As String, sArq As String, bOK As Boolean
   sPastaDL = appPathApp & "IBPT_Auto"
   On Error Resume Next
   MkDir sPastaDL
   On Error GoTo 0

   Me.MousePointer = 11
   mensagemErro = ""
   sArq = GoogleBaixarArquivo(sPastaDL)
   Me.MousePointer = 0

   If Trim(sArq) = "" Then
      MsgBox "N" & Chr(227) & "o foi poss" & Chr(237) & "vel baixar a tabela IBPT da nuvem." & _
             IIf(mensagemErro <> "", vbCrLf & vbCrLf & mensagemErro, ""), vbCritical, "Tabela IBPT"
      Exit Sub
   End If

   frmImportarIBPT.Show
   frmImportarIBPT.lblProgresso.Visible = True
   DoEvents
   Me.MousePointer = 11
   bOK = frmImportarIBPT.ImportarIBPTdeArquivo(sPastaDL & "\\" & sArq, False)
   Me.MousePointer = 0

   On Error Resume Next
   Kill sPastaDL & "\\" & sArq
End Sub''',
b'''Private Sub Menu_FISCAL_Reforma_IBPTNuvem_Click()
   ' baixa a tabela IBPT mais recente do Shared Drive "IBPT" e importa na hora (sem gate de data).
   ' frmImportarIBPT nunca aparece: o andamento vai todo pra frmIBPTProgresso (barra).
   If MsgBox("Baixar a tabela IBPT mais recente da nuvem e atualizar agora?", vbQuestion + vbYesNo, "Tabela IBPT") <> vbYes Then Exit Sub

   Dim sPastaDL As String, sArq As String, bOK As Boolean
   sPastaDL = appPathApp & "IBPT_Auto"
   On Error Resume Next
   MkDir sPastaDL
   On Error GoTo 0

   Me.MousePointer = 11
   frmIBPTProgresso.Iniciar "Baixando a tabela IBPT da nuvem..." & vbCrLf & _
      "Isso pode levar alguns minutos. N" & Chr(227) & "o feche o sistema."
   mensagemErro = ""
   sArq = GoogleBaixarArquivo(sPastaDL)

   If Trim(sArq) = "" Then
      Unload frmIBPTProgresso
      Me.MousePointer = 0
      MsgBox "N" & Chr(227) & "o foi poss" & Chr(237) & "vel baixar a tabela IBPT da nuvem." & _
             IIf(mensagemErro <> "", vbCrLf & vbCrLf & mensagemErro, ""), vbCritical, "Tabela IBPT"
      Exit Sub
   End If

   frmIBPTProgresso.DefinirMensagem "Importando a tabela IBPT..." & vbCrLf & _
      "Aguarde, n" & Chr(227) & "o feche o sistema."
   bOK = frmImportarIBPT.ImportarIBPTdeArquivo(sPastaDL & "\\" & sArq, True, frmIBPTProgresso)
   Unload frmImportarIBPT
   Unload frmIBPTProgresso
   Me.MousePointer = 0

   On Error Resume Next
   Kill sPastaDL & "\\" & sArq
   On Error GoTo 0

   If bOK Then
      MsgBox "Tabela IBPT atualizada com sucesso!", vbInformation, "Tabela IBPT"
   Else
      MsgBox "Falha ao importar a tabela IBPT." & _
             IIf(mensagemErro <> "", vbCrLf & vbCrLf & mensagemErro, ""), vbCritical, "Tabela IBPT"
   End If
End Sub''',
    ),
])
print("Feito.")
