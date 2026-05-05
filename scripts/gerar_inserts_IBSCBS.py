# -*- coding: utf-8 -*-
import json

def esc(v):
    if v is None or v == '':
        return 'NULL'
    return "'" + str(v).replace("'", "''") + "'"

def esc_nn(v):
    """esc para colunas NOT NULL — retorna '' em vez de NULL"""
    if v is None or v == '':
        return "''"
    return "'" + str(v).replace("'", "''") + "'"

def bit(v):
    return '1' if str(v) == '1' else '0'

def dec(v):
    try:
        return str(float(v))
    except:
        return '0'

def dt(v):
    if not v or v == '':
        return 'NULL'
    return "CONVERT(date, '" + str(v) + "', 23)"

def dt_nn(v):
    """dt para colunas NOT NULL — retorna data mínima em vez de NULL"""
    if not v or v == '':
        return "CONVERT(date, '1900-01-01', 23)"
    return "CONVERT(date, '" + str(v) + "', 23)"

header = (
    "-- ============================================================\n"
    "-- Script gerado automaticamente em {data}\n"
    "-- Execute no banco de dados de destino\n"
    "-- ============================================================\n"
    "SET NOCOUNT ON;\n"
    "SET ANSI_NULLS ON;\n"
    "SET QUOTED_IDENTIFIER ON;\n"
    "GO\n\n"
)

from datetime import date
hoje = date.today().strftime('%d/%m/%Y')

# ── TbIBSCBS ──────────────────────────────────────────────────────────────────
with open('C:/Projeto/TbIBSCBS.txt', encoding='utf-8') as f:
    data1 = json.load(f)

lines1 = [header.format(data=hoje)]
lines1.append('-- Limpa dados anteriores para evitar duplicidade')
lines1.append('TRUNCATE TABLE TbIBSCBSClassTrib; -- filho primeiro (FK)')
lines1.append('DELETE FROM TbIBSCBS; -- TRUNCATE nao permitido em tabela pai com FK')
lines1.append('GO\n')

lines1.append('INSERT INTO TbIBSCBS (CST, DescricaoIBSCBS, ind_gIBSCBS, ind_gIBSCBSMono, ind_gRed, ind_gDif, ind_gTransfCred, ind_gCredPresIBSZFM, ind_gAjusteCompet, ind_RedutorBC)')
lines1.append('VALUES')
rows1 = []
for r in data1:
    rows1.append("  ({},{},{},{},{},{},{},{},{},{})".format(
        esc(r['CST']),
        esc(r['Descricao']),
        bit(r['ind_gIBSCBS']),
        bit(r['ind_gIBSCBSMono']),
        bit(r['ind_gRed']),
        bit(r['ind_gDif']),
        bit(r['ind_gTransfCred']),
        bit(r['ind_gCredPresIBSZFM']),
        bit(r['ind_gAjusteCompet']),
        bit(r['ind_RedutorBC'])
    ))
lines1.append(',\n'.join(rows1) + ';')
lines1.append('GO')
lines1.append('\nPRINT \'TbIBSCBS: \' + CAST(@@ROWCOUNT AS VARCHAR) + \' registros inseridos.\'')
lines1.append('GO')

with open('C:/Projeto/scripts/Preencher_TbIBSCBS.sql', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines1))
print('Preencher_TbIBSCBS.sql gerado ({} registros)'.format(len(data1)))

# ── TbIBSCBSClassTrib ─────────────────────────────────────────────────────────
with open('C:/Projeto/TbIBSCBSClassTrib.txt', encoding='utf-8') as f:
    data2 = json.load(f)

cols = ('CST, DescricaoIBSCBS, cClassTrib, NomecClassTrib, DescricaocClassTrib, '
        'LC_Redacao, LC_214_25, TipoDeAliquota, pRedIBS, pRedCBS, '
        'ind_gTribRegular, ind_gCredPresOper, ind_gMonoPadrao, indMonoReten, indMonoRet, indMonoDif, '
        'Credito_para, dIniVig, dFimVig, DataAtualizacao, ind_gEstornoCred, '
        'indNFeABI, indNFe, indNFCe, indCTe, indCTeOS, indBPe, indBPeTA, indBPeTM, '
        'indNF3e, indNFSe, indNFSe_Via, indNFCom, indNFAg, indNFGas, indDERE, Anexo, Link')

lines2 = [header.format(data=hoje)]
lines2.append('-- Limpa dados anteriores para evitar duplicidade')
lines2.append('TRUNCATE TABLE TbIBSCBSClassTrib;')
lines2.append('GO\n')

batch_size = 100
batches = [data2[i:i+batch_size] for i in range(0, len(data2), batch_size)]
total = 0

for idx, b in enumerate(batches):
    rows2 = []
    for r in b:
        rows2.append("  ({},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{})".format(
            esc(r['CST']),
            esc(r['Descricao_CST']),
            esc(r['cClassTrib']),
            esc(r['Nome_cClassTrib']),
            esc(r['Descricao_cClassTrib']),
            esc_nn(r.get('LC_Redacao','')),
            esc_nn(r.get('LC_214_25','')),
            esc(r.get('TipoDeAliquota','')),
            dec(r.get('pRedIBS',0)),
            dec(r.get('pRedCBS',0)),
            bit(r.get('ind_gTribRegular',0)),
            bit(r.get('ind_gCredPresOper',0)),
            bit(r.get('ind_gMonoPadrao',0)),
            bit(r.get('indMonoReten',0)),
            bit(r.get('indMonoRet',0)),
            bit(r.get('indMonoDif',0)),
            esc_nn(r.get('Credito_para','')),
            dt(r.get('dIniVig','')),
            dt_nn(r.get('dFimVig','')),
            dt(r.get('DataAtualizacao','')),
            bit(r.get('ind_gEstornoCred',0)),
            bit(r.get('indNFeABI',0)),
            bit(r.get('indNFe',0)),
            bit(r.get('indNFCe',0)),
            bit(r.get('indCTe',0)),
            bit(r.get('indCTeOS',0)),
            bit(r.get('indBPe',0)),
            bit(r.get('indBPeTA',0)),
            bit(r.get('indBPeTM',0)),
            bit(r.get('indNF3e',0)),
            bit(r.get('indNFSe',0)),
            bit(r.get('indNFSe_Via',0)),
            bit(r.get('indNFCom',0)),
            bit(r.get('indNFAg',0)),
            bit(r.get('indNFGas',0)),
            bit(r.get('indDERE',0)),
            esc_nn(r.get('ANEXO', r.get('Anexo',''))),
            esc(r.get('Link',''))
        ))
    lines2.append('-- Lote {} de {} ({} registros)'.format(idx+1, len(batches), len(b)))
    lines2.append('INSERT INTO TbIBSCBSClassTrib ({})'.format(cols))
    lines2.append('VALUES')
    lines2.append(',\n'.join(rows2) + ';')
    lines2.append('GO\n')
    total += len(b)

lines2.append("PRINT 'TbIBSCBSClassTrib: {} registros inseridos.'".format(total))
lines2.append('GO')

with open('C:/Projeto/scripts/Preencher_TbIBSCBSClassTrib.sql', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines2))
print('Preencher_TbIBSCBSClassTrib.sql gerado ({} registros em {} lotes)'.format(total, len(batches)))
