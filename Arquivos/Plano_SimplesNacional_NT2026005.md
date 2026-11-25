# Plano: Simples Nacional na NF-e/NFC-e (NT 2026.005 v1.00)

Criado em 24/11/2026. **Status: plano, nada implementado.** Depende de respostas do dev da DLL (seção 7).

Fontes: `Arquivos/NFe_NT_2026.005_v1.00_RTC_SimplesNacional.pdf`, `OnlineCommerce/NFE/docs/ROTEIRO_ATUALIZACOES_NFE_NFCE.md` (seção "Simples Nacional — NT 2026.005") e leitura do `Compartilhado/Modulos/modNFe.bas`.

## 1. O que a NT exige

Prazo da NT: campos novos em teste até 26/10/2026 e produção em 16/11/2026. Um Aviso de 12/11/2026 (fontes de terceiros: NFE.io, Inventti, Senior; texto oficial não localizado, confirmar no Portal Nacional da NF-e) adiou para **14/12/2026** a produção das regras de rejeição; os campos novos já podem ser enviados desde 16/11. O evento 412140 fica para implementação futura.

| Emitente | `tpRegimeIBSCBS` (C21a) | `cAtivSN` (I17d, por item) | `gIBSCBS` | `gTribSN` (UB115a) |
|---|---|---|---|---|
| CRT 1 ou 2, **com** regime regular IBS/CBS | `1` | informar | informar (UF, Mun, CBS) | não informar |
| CRT 1, **sem** regime regular | não informar | informar | **não informar** | informar `gIBSSN` + `gCBSSN` |
| CRT 2 (excesso sublimite), sem regime regular | não informar | informar | informar com CBS de alíquota zero | informar só `gCBSSN` |
| CRT 3 com opção pendente pelo CRT 1, sem regime regular | `2` | informar | informar | informar `gIBSSN` + `gCBSSN` |
| CRT 3 com opção pendente pelo CRT 2, sem regime regular | `3` | informar | informar | informar só `gCBSSN` |
| CRT 3 com opção pendente (CRT 1 ou 2), com regime regular | `4` | informar | informar | não informar |
| CRT 3 normal | não informar | não informar | informar | não informar |
| MEI (CRT 4) | não informar | não informar | **não informar** | não informar |

A tabela acima foi conferida com as tabelas 2.1 e 2.3 da NT. Uma resposta de outra IA (imagens 293-295) resumiu os casos 2 e 3 do Regime Normal pendente como só `gIBSCBS`; a NT manda também `gTribSN` (gIBSSN+gCBSSN no 2, só gCBSSN no 3). Vale a NT.

Regras adicionais:
- `cAtivSN`: 01 revenda/indústria Anexo I; 02 demais bens sem ICMS Anexo I; 03 e 04 ZFM Anexo II; 90 não tributável.
- `gCBSSN` é obrigatório dentro do `gTribSN`; `gIBSSN` é opcional.
- CST 620 (combustível monofásico): preencher `gIBSCBSMono` e **não** preencher `gTribSN`.
- Totais: `total/IBSCBSTot/gTribSN` = soma dos itens (rejeições 1275, 1276, 1277, 1284). A DLL soma sozinha no `GerarXML`.
- Devolução (`finNFe=4`) e notas de crédito/débito por recusa: o grupo IBSCBS do item é **espelho** do item da nota original.

## 2. Situação atual do código (divergências)

`modNFe.bas`, NF-e e NFC-e, comentário `' IBS/CBS: todos os regimes`:

| Ponto | Linhas | O que faz hoje |
|---|---|---|
| NF-e, item não combustível | ~805-828 | `GerarItensImpostoIBSCBS` para qualquer CRT |
| NF-e, item combustível | ~844-867 | idem (bloco duplicado) |
| NFC-e, item | ~2066-2089 | idem |
| NF-e, total | ~914-916 | `GerarTotalIBSCBS` se `TotvBCCBSIBS > 0` |
| NFC-e, total | ~2152 | idem |
| Emitente | 517 (NF-e), 1904 (NFC-e) | `GerarEmitente`, sem `GerarEmitenteRegimeIBSCBS` |

Valores de IBS/CBS por item são calculados como Regime Normal (`NFe_Completa.frm` ~4600 e `modNFe.bas` ~1785), sem olhar o CRT.

Consequências: CRT 1 sem regime regular envia `gIBSCBS` (a NT manda não enviar) e não envia `cAtivSN` nem `gTribSN`. MEI envia grupo IBSCBS (a NT manda não enviar). Ninguém envia `tpRegimeIBSCBS`.

## 3. Dados novos

### 3.1 Empresa
- `Empresa.OptaRegimeRegularIBSCBS` bit NOT NULL DEFAULT 0 (Simples que optou pelo regime regular).
- `Empresa.SimplesOpcaoPendente` tinyint NOT NULL DEFAULT 0 (0 = não; 1 = pendente pelo CRT 1; 2 = pendente pelo CRT 2). Só vale com CRT = 3.
- `tpRegimeIBSCBS` é **derivado** (função), não gravado:
  - CRT 1 ou 2 e `OptaRegimeRegular = 1` → 1
  - CRT 3 e pendente = 1 e `OptaRegimeRegular = 0` → 2
  - CRT 3 e pendente = 2 e `OptaRegimeRegular = 0` → 3
  - CRT 3 e pendente > 0 e `OptaRegimeRegular = 1` → 4
  - demais → 0 (não chamar o método)
- Tela: `Configuracao_Geral.frm` ou o cadastro da empresa. Verificar onde o CRT é editado hoje. Desenho decidido em 24/11/2026:
  - Com CRT 1 ou 2: seletor "Apuração de IBS/CBS: Pelo Simples Nacional (DAS) [padrão]" (grava `OptaRegimeRegular = 0`) ou "Regime Regular (fora do DAS)" (grava 1).
  - Com CRT 3: seção "Situações especiais / sub judice" com a opção pendente (nenhuma / pelo CRT 1 / pelo CRT 2) e o mesmo seletor de regime regular. Só esses casos geram `tpRegimeIBSCBS` 2, 3 ou 4.
  - Com CRT 4 (MEI): nada a escolher; o grupo IBSCBS não é enviado.
  - Esconder termos técnicos da SEFAZ do usuário final; mostrar o texto em português simples.

### 3.2 Atividade (`cAtivSN`)
**Decidido (24/11/2026): por produto, com padrão na empresa.** A mesma nota pode ter revenda (01) e item não tributável (90) ou sem ICMS (02); por isso o valor não pode ser fixo na empresa.
- `Produtos.cAtivSN` char(2) NULL: valor do item; o usuário pode alterar no cadastro do produto.
- `Empresa.cAtivSNPadrao` char(2) DEFAULT '01': só pré-preenche produto novo e serve de fallback quando o produto estiver sem valor.
- Sugestão automática (opcional): 90 quando o CST de IBS/CBS ou o CFOP indicar não tributável; 02 quando o CSOSN indicar sem incidência de ICMS; senão 01. Sempre editável.
- Gravar o valor usado no item da nota: `NotaFiscalItens.cAtivSN` e `TbNFCe_Itens.cAtivSN` (char(2) NULL), para a reemissão e a devolução serem fiéis.

### 3.3 Item (gTribSN)
Colunas nos itens para os valores calculados, nas duas tabelas de itens:
- `pIBSSN`, `vIBSSN`, `pCBSSN`, `vCBSSN` decimal(13,4/2) NULL.
- Totais no cabeçalho (`NotaFiscal` e a tabela de cabeçalho da NFC-e): `vIBSSNTot`, `vCBSSNTot`.

### 3.4 Script de manifesto
Um script novo `Alter_SimplesNacional_NT2026005.sql` (GERAL), idempotente (`COL_LENGTH ... IS NULL`), próximo número livre do manifesto. Um só script para Empresa, Produtos, NotaFiscalItens, TbNFCe_Itens e os cabeçalhos.

## 4. Mudanças no código

### 4.1 Função auxiliar (módulo novo `modSimplesNacional.bas` em `Compartilhado/Modulos`)
- `TipoRegimeIBSCBS(rEmpresa) As Integer` → 0..4 conforme 3.1.
- `UsaTribSN(CRT, tpRegime) As Boolean` → True quando CRT 1 ou 2 sem regime regular (e CRT 3 com `tpRegime` 2 ou 3).
- `InformaGIBSCBSComTribSN(...)`: casos em que `gIBSCBS` e `gTribSN` coexistem (tpRegime 2 e 3).
- `cAtivSNDoItem(codProduto)`.
- Centralizar a regra aqui para NF-e, NFC-e e PDV usarem a mesma.

### 4.2 `modNFe.bas`
1. **Emitente** (linhas 517 e 1904): depois do `GerarEmitente`, `If TipoRegimeIBSCBS > 0 Then sistNFe.GerarEmitenteRegimeIBSCBS(TipoRegimeIBSCBS, mensagemAlerta, mensagemErro)`.
2. **Itens**, nos 3 blocos (não combustível, combustível, NFC-e):
   - Se CRT = 4 (MEI): pular todo o grupo IBSCBS.
   - Se CRT 1 ou 2, ou CRT 3 com `tpRegime` 2, 3 ou 4: chamar `GerarItensAtividadeSimplesNacional(cAtivSN, ...)`.
   - Se CRT 1 sem regime regular: **não** chamar `GerarItensImpostoIBSCBS`; chamar `GerarItensImpostoSimplesNacional(cst, cClassTrib, informarIBS, pIBSSN, vIBSSN, pCBSSN, vCBSSN, ...)`.
   - Se CRT 2 sem regime regular: chamar o `IBSCBS` com CBS de alíquota zero + `GerarItensImpostoSimplesNacional` só com CBS (`informarIBS=False`).
   - Se `tpRegime` 2 ou 3: `GerarItensImpostoIBSCBS` **e** `GerarItensImpostoSimplesNacional`.
   - Se `tpRegime` 1 ou 4: só `GerarItensImpostoIBSCBS` (com regime regular).
   - Item combustível CST 620: não chamar `GerarItensImpostoSimplesNacional`.
3. **Totais** (linhas 915 e 2152): para CRT 1 sem regime regular não chamar `GerarTotalIBSCBS` com valores de Regime Normal. Confirmar com o dev se `GerarTotalIBSCBS` é necessário mesmo assim (seção 7).
4. **Devolução**: copiar os campos IBSCBS/SN do item original (já há conferência de `nfeProc` na DLL).
5. Os três blocos são hoje cópias; extrair uma `Sub GerarIBSCBSItem(...)` evita triplicar a nova lógica. Fazer isso **antes** das mudanças de regra, em commit separado, para não misturar refatoração com comportamento.

### 4.3 Cálculo dos valores SN
`NFe_Completa.frm` (~4600) e `modNFe.bas` (~1785):
- CRT 1 ou 2 sem regime regular: não calcular IBS UF/Mun/CBS de Regime Normal. Calcular `vIBSSN` e `vCBSSN` pela alíquota efetiva do Simples (seção 7, pergunta 2). `IBS_vBC` pode continuar sendo gravado para relatórios.
- Relatórios e apuração que somam `IBS_vIBSUF` etc. (`NFe_Completa.frm` ~5384, ~9288; `NFCe_Consultar_Produtos.frm` ~1518) hoje mostram Regime Normal. Revisar para não exibir valores que não foram enviados.

### 4.4 PDV (NFC-e)
`PDV/Forms/*` que preenchem `TbNFCe_Itens`: gravar `cAtivSN` e os campos SN no item no momento da venda (mesma função auxiliar).

### 4.5 Demais telas
- Cadastro de produto: campo `cAtivSN`, visível só para empresa do Simples (CRT 1 ou 2, ou CRT 3 com opção pendente).
- Alerta ao emitir se faltar `cAtivSN` ou o enquadramento (`OptaRegimeRegular`) em empresa do Simples.

## 5. Ordem de implementação

1. Script SQL + campos na Empresa e telas de configuração.
2. Módulo `modSimplesNacional.bas` com a lógica de decisão (sem uso ainda).
3. Refatorar os 3 blocos IBS/CBS de `modNFe.bas` numa rotina única (sem mudar comportamento; confirmar com nota de teste CRT 3).
4. Aplicar a regra do Simples dentro da rotina única, atrás de `CRT <> 3 Or tpRegime > 0`, para CRT 3 normal ficar byte a byte igual.
5. Cálculo dos valores SN (depende da pergunta 2).
6. Devolução/crédito: espelhar o item original.
7. PDV/NFC-e.
8. Relatórios e apuração.
9. Testes em homologação (seção 6) e só então manifesto/versão.

## 6. Testes (homologação, NF-e e NFC-e)

1. CRT 3 normal: XML idêntico ao de hoje (comparar com nota anterior).
2. CRT 1 sem regime regular: sem `gIBSCBS`; com `cAtivSN`, `gTribSN` (IBS e CBS); total `gTribSN` = soma.
3. CRT 1 com regime regular: `tpRegimeIBSCBS=1`, `cAtivSN`, `gIBSCBS`, sem `gTribSN`.
4. CRT 2 sem regime regular: `gIBSCBS` com CBS zero e `gTribSN` só CBS.
5. CRT 3 pendente (valores 2, 3, 4).
6. MEI: sem grupo IBSCBS.
7. Item combustível CST 620: sem `gTribSN`.
8. Devolução de nota do Simples.
9. Rejeições esperadas: 1275, 1276, 1277, 1284 (forçar divergência para ver a mensagem).

## 7. Perguntas ao dev da DLL (pendentes)

1. Com a DLL nova, o `GerarXML` rejeita (ou só avisa) quando faltam `cAtivSN` e `gTribSN` em empresa do Simples? E quando `gIBSCBS` vai junto com `gTribSN` onde a NT manda não enviar?
2. Quem calcula `pIBSSN`/`pCBSSN` (alíquota efetiva do Simples, que depende de RBT12 e faixa)? A DLL, o ERP, ou vem de tabela do Fisco? O roteiro diz que valores fiscais ficam a cargo do chamador.
3. `GerarTotalIBSCBS` ainda deve ser chamado para CRT 1 sem regime regular, ou o total vem só do `gTribSN`?
4. ~~A SEFAZ já rejeita?~~ Respondido: rejeições só em produção a partir de 14/12/2026 (confirmar no Portal).
5. Qual `cAtivSN` usar para serviço (OS/NFS-e)? A NT só lista valores de bens materiais e 90.

### Respostas da outra IA (genéricas, não conhecem a snfe.dll)

- Q1: disse que a DLL só valida XSD. **Incorreto para esta DLL**: o ROTEIRO lista dezenas de regras de negócio validadas no `GerarXML`. Segue aberta (regras 1275 a 1277 e 1284 estão incluídas?).
- Q2: o ERP calcula a alíquota efetiva pela fórmula do PGDAS-D `(RBT12 x alíquota nominal - parcela a deduzir) / RBT12`, parte de IBS/CBS por anexo. Resoluções CGSN citadas não conferidas: **confirmar com contador/dev**.
- Q3: CRT 1 sem regime regular não chama `GerarTotalIBSCBS` (total em W59h a W59j); CRT 2 chama os dois totalizadores. Bate com a NT. Falta confirmar se a DLL soma o `gTribSN` sozinha.
- Q5: NFS-e não usa `cAtivSN`; NF-e com serviço conjugado usar 90. Não confirmado, fica para o dev.

### Impacto: tela de alíquotas efetivas do Simples
Não existe nada hoje. Será preciso tabela/tela de parâmetros (alíquota efetiva de IBS e de CBS por anexo e competência, preenchida pelo contador ou calculada pelo RBT12). É o maior trabalho restante; decisão de como fazer é do usuário.

## 8. Decisões do usuário

- ~~`cAtivSN` por empresa ou por produto~~: decidido por produto, com padrão na empresa (seção 3.2).
- Quais clientes são do Simples com regime regular de IBS/CBS (campo `OptaRegimeRegular`)? Provavelmente nenhum hoje; manter DEFAULT 0.
- Plano de rollout: atualizar clientes do Simples primeiro, depois o resto (CRT 3 não é afetado se a condição do passo 4 da seção 5 for respeitada).

## 9. Riscos

- Mudar o bloco IBS/CBS sem proteção atinge todos os clientes, inclusive o CRT 3. Proteger por condição e comparar XML antes/depois.
- Edição de `modNFe.bas` e `NFe_Completa.frm`: arquivos Windows-1252, grandes; editar via Python binário e conferir com `git diff`.
- A DLL, a TLB 5.3 e os 3 `.vbp` (referência `#5.3#`) estão locais, sem commit; a outra máquina e os clientes precisam receber DLL e TLB junto com o código novo.
- Sem alíquota efetiva correta do Simples, os valores de `gTribSN` saem errados; é preferível não enviar do que enviar valor inventado, mas a NT exige o grupo (`gCBSSN` obrigatório).
