# -*- coding: utf-8 -*-
"""
Monta o novo arquivo OrdemServico\Forms\OS_Consulta.frm a partir dos
blocos extraidos de OS_Recapadora.frm (scripts\extract_consulta_parts.py).
Grava em cp1252 com CRLF, como todo .frm do projeto.
"""

SCRATCH = r"C:\Users\NOTEBOOK\AppData\Local\Temp\claude\C--projeto\916fb1c0-4fd5-437b-8d03-a83de36ec5b2\scratchpad"
OUT_PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"


def load(name):
    with open(f"{SCRATCH}\\consulta_{name}.txt", "r", encoding="cp1252") as f:
        return f.read()


def reposition_left(block, old_left, new_left):
    old_line = f"         Left            =   {old_left}"
    new_line = f"         Left            =   {new_left}"
    n = block.count(old_line)
    assert n == 1, f"esperado 1 ocorrencia de {old_line!r} em bloco, achou {n}"
    return block.replace(old_line, new_line, 1)


frame2 = load("Frame2")
frame2 = reposition_left(frame2, "-74880", "120")

grid = load("Grid")
grid = reposition_left(grid, "-74880", "120")

lblquant = load("lblQuant")
lblquant = reposition_left(lblquant, "-74880", "120")

lbltotal = load("lblTotalConsulta")
lbltotal = reposition_left(lbltotal, "-62760", "12240")

header = '''VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form OS_Consulta
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Ordens de Serviço"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar
      Caption         =   "Fechar"
      Height          =   375
      Left            =   120
      TabIndex        =   205
      Top             =   9060
      Width           =   1500
   End
'''

footer_controls_end = "End\r\n"

attributes = '''Attribute VB_Name = "OS_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lCodOSSelecionado As Long
Private moCombo As cComboHelper
Dim printSQL As String

Private Sub Form_Load()
Set moCombo = New cComboHelper
lCodOSSelecionado = 0

Preencher_TipoServico
Preencher_Mostrar
Preencher_Status
Preencher_Criterios
Preencher_Indice

cboConsultaMostrar.ListIndex = 0
cboConsultaStatus.ListIndex = 0
cboConsultaCriterios.ListIndex = 0
cboTipoServico.ListIndex = 0
cboIndice.ListIndex = 0

MostrarGrid_OS
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
lCodOSSelecionado = 0
Me.Hide
End Sub

Private Sub cmdFechar_Click()
lCodOSSelecionado = 0
Me.Hide
End Sub

'''

# Grid_DblClick adaptado (sem tocar em SSTab1/frmSecundario/etc de OS_Recapadora)
grid_dblclick_novo = '''Private Sub Grid_DblClick()
lCodOSSelecionado = Val(Grid.TextMatrix(Grid.Row, 0))
Me.Hide
End Sub

'''

soma_grid_novo = '''Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer
Dim Valor As Currency

Valor = 0
For i = 0 To var_Grid.Rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function

'''

code_sub_order = [
    "cboConsultaCriterios_Click",
    "cboConsultaCriterios_GotFocus",
    "cboConsultaCriterios_Validate",
    "cboConsultaMostrar_Change",
    "cboConsultaMostrar_Click",
    "cboConsultaMostrar_GotFocus",
    "cboConsultaMostrar_Validate",
    "cboConsultaStatus_Change",
    "cboConsultaStatus_Click",
    "cboConsultaStatus_GotFocus",
    "cboConsultaStatus_Validate",
    "cboIndice_Change",
    "cboIndice_Click",
    "cboIndice_GotFocus",
    "cboLocalizar_GotFocus",
    "cboLocalizar_LostFocus",
    "cboTipoServico_Change",
    "cboTipoServico_Click",
    "cboTipoServico_GotFocus",
    "cmdExibir_Click",
    "cmdImprimirConsulta_Click",
]

code_parts = [attributes]
for name in code_sub_order:
    code_parts.append(load(name))
    code_parts.append("\r\n\r\n")

code_parts.append(grid_dblclick_novo)

for name in ["MostrarGrid_OS", "FormatarGrid_OS", "Preencher_Criterios", "Preencher_Indice",
             "Preencher_Mostrar", "Preencher_Status", "Preencher_TipoServico"]:
    code_parts.append(load(name))
    code_parts.append("\r\n\r\n")

code_parts.append(soma_grid_novo)

full_text = (
    header
    + frame2 + "\r\n"
    + grid + "\r\n"
    + lblquant + "\r\n"
    + lbltotal + "\r\n"
    + footer_controls_end
    + "".join(code_parts)
)

# normaliza quebras de linha
full_text = full_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")

with open(OUT_PATH, "wb") as f:
    f.write(full_text.encode("cp1252"))

print("OK - OS_Consulta.frm criado,", len(full_text.split("\r\n")), "linhas")
