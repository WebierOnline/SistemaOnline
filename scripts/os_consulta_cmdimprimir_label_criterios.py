# -*- coding: utf-8 -*-
"""
cmdImprimir_Click (OS_Consulta.frm):
- Label4.Caption do REL_OS_Consulta passa a depender de vTipoOS
  (Automoveis/Motocicletas/Recapadora -> "CLIENTE / VEICULOS";
   demais -> "CLIENTE / EQUIPAMENTO", mesmo agrupamento ja usado
   nas 19 queries de MostrarGrid_OS/_Refinado).
- Preenche rfCons1 (tipo de consulta: Simples/Refinado + criterio),
  rfCons2 (valor do criterio usado) e rfCons3 (filtros secundarios:
  status, situacao, pagamento, tipo de servico), tornando visiveis
  ReportField9, rfCons1 e ReportField10 (ocultos por padrao no design).
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


anchor = find_line_exact("REL_OS_Consulta.Relatorio.NomeImpressora = var_Impressora")

new_lines = [
    "If vTipoOS = \"Autom\xf3veis\" Or vTipoOS = \"Motocicletas\" Or vTipoOS = \"Recapadora\" Then",
    "   REL_OS_Consulta.Label4.Caption = \"CLIENTE / VE\xcdCULOS\"",
    "Else",
    "   REL_OS_Consulta.Label4.Caption = \"CLIENTE / EQUIPAMENTO\"",
    "End If",
    "",
    "Dim sTipoConsulta As String",
    "Dim sCriterio As String",
    "Dim sFiltrosAdic As String",
    "",
    "If optFiltroRefinado.Value = True Then",
    "   sTipoConsulta = \"Refinado - \" & IIf(chkPlaca.Value = 1, \"PLACA\", \"CHASSI\")",
    "   sCriterio = txtFiltroRefinado.Text",
    "Else",
    "   sTipoConsulta = \"Simples - \" & cboConsultaCriterios.Text",
    "   Select Case cboConsultaCriterios.Text",
    "      Case \"CLIENTE\"",
    "         sCriterio = txtCodClienteLocalizar.Text",
    "      Case \"C\xd3D. OS\"",
    "         sCriterio = cboLocalizar.Text",
    "      Case \"DATA\"",
    "         sCriterio = mskDataConsulta.Text",
    "      Case \"PER\xcdODO\"",
    "         sCriterio = mskPeriodoInicio.Text & \" a \" & mskPeriodoFim.Text",
    "      Case \"MENSAL\"",
    "         sCriterio = cboMesConsulta.Text & \"/\" & cboAnoConsulta.Text",
    "      Case Else",
    "         sCriterio = \"TODOS\"",
    "   End Select",
    "End If",
    "",
    "sFiltrosAdic = \"Status: \" & cboConsultaStatus.Text & \"  |  Situa\xe7\xe3o: \" & cboConsultaMostrar.Text & \"  |  Pagamento: \" & cboTipo.Text & \"  |  Tipo Servi\xe7o: \" & cboTipoServico.Text",
    "",
    "REL_OS_Consulta.ReportField9.Visible = True",
    "REL_OS_Consulta.ReportField10.Visible = True",
    "REL_OS_Consulta.rfCons1.Visible = True",
    "REL_OS_Consulta.rfCons1.Caption = sTipoConsulta",
    "REL_OS_Consulta.rfCons2.Caption = sCriterio",
    "REL_OS_Consulta.rfCons3.Caption = sFiltrosAdic",
    "",
]

lines[anchor:anchor] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - {len(new_lines)} linhas inseridas em cmdImprimir_Click antes da linha {anchor}")
