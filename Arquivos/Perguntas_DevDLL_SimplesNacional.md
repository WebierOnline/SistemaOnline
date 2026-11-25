# Perguntas sobre a snfe.dll (v5.3) - Simples Nacional na NF-e/NFC-e (NT 2026.005)

Este documento e para ser lido por quem conhece o codigo da `snfe.dll` (o desenvolvedor ou a IA dele com acesso ao fonte).
Responda cada pergunta com **SIM / NAO / PARCIAL** na primeira linha e, logo abaixo, uma explicacao curta com o **nome do metodo, da classe ou da regra** no codigo que comprova a resposta. Se nao souber, escreva **NAO SEI** em vez de supor.

## Contexto

- Sistema cliente: VB6, chama a DLL por COM (`snfe.Util`). A TLB e a 5.3.
- Metodos novos que o sistema vai usar:
  - `GerarEmitenteRegimeIBSCBS(tpRegimeIBSCBS)` com valores 1 a 4 (tag C21a).
  - `GerarItensAtividadeSimplesNacional(cAtivSN)` (tag I17d, valores 01, 02, 03, 04 e 90).
  - `GerarItensImpostoSimplesNacional(cst, cClassTrib, informarIBS, pIBSSN, vIBSSN, pCBSSN, vCBSSN)` (grupo UB115a, gTribSN).
  - `GerarTotalIBSCBS(...)` (grupo W03, totais IBS/CBS).
- O **sistema calcula** `pIBSSN`, `vIBSSN`, `pCBSSN` e `vCBSSN` (alíquota efetiva do Simples) e passa prontos para a DLL. Isso ja esta decidido; nao e pergunta.
- Regras da NT 2026.005 v1.00 que importam aqui:
  - W59h-10 (rejeicao 1275): ha item com `gTribSN` e o total `gTribSN` (W59h) nao foi informado.
  - W59h-20 (1276): total `gTribSN` informado, mas nenhum item tem `gTribSN`.
  - W59i-10 (1277): `vIBSSN` total (W59i) difere da soma de `vIBSSN` dos itens.
  - W59j-10 (1284): `vCBSSN` total (W59j) difere da soma de `vCBSSN` dos itens.
  - CRT 1 sem regime regular de IBS/CBS: informar `cAtivSN` e `gTribSN` (IBS e CBS), **sem** `gIBSCBS`.
  - CRT 2 sem regime regular: `cAtivSN`, `gIBSCBS` com CBS zerada e `gTribSN` so com CBS.
  - CRT 1 ou 2 com regime regular: `tpRegimeIBSCBS=1`, `cAtivSN`, `gIBSCBS`, sem `gTribSN`.
  - CRT 3 com opcao pendente pelo CRT 1 ou 2: `tpRegimeIBSCBS` 2, 3 ou 4 conforme a NT.
  - CRT 3 normal e MEI: sem alteracao (MEI nao informa `gIBSCBS`).
  - CST 620 (combustivel monofasico): informar `gIBSCBSMono`, sem `gTribSN`.
  - Compra governamental em 2027 e 2028: se `tpEnteGov` = 1 (Uniao), `pIBSSN` = 0 e `pCBSSN` = soma das aliquotas efetivas de IBS e CBS; se for outro ente, cada um com a sua.

## Perguntas

### 1. Total do Simples (campos W59h, W59i e W59j)
Quando os itens trazem `gTribSN` (via `GerarItensImpostoSimplesNacional`), o `GerarXML` monta **sozinho** o grupo total `gTribSN` (W59h) com a soma de `vIBSSN` (W59i) e de `vCBSSN` (W59j) dos itens?
Se NAO, qual metodo o sistema chama para informar esse total, e com quais parametros?

### 2. `GerarTotalIBSCBS` para empresa do Simples sem regime regular (CRT 1)
Nesse caso o sistema deve chamar `GerarTotalIBSCBS`, ou nao chamar (porque o total vem so do `gTribSN`)?
O que acontece (erro, XML invalido ou ignorado) se chamar mesmo assim? E se **nao** chamar, mas houver `gTribSN` nos itens?
Para CRT 2 sem regime regular (que tem `gIBSCBS` com CBS zerada e `gTribSN` so com CBS), quais totalizadores devem ser chamados?

### 3. A DLL bloqueia nota errada no `GerarXML`?
Para cada caso abaixo, informe se o `GerarXML` **recusa a nota** (erro, nota nao e gerada) ou **deixa passar** para a SEFAZ validar:
- a) Empresa do Simples (CRT 1 ou 2) sem `cAtivSN` em algum item.
- b) Empresa do Simples sem `gTribSN` onde a NT exige.
- c) `gIBSCBS` enviado em CRT 1 sem regime regular (a NT manda nao enviar).
- d) Total do Simples (W59i, W59j) diferente da soma dos itens (rejeicoes 1277 e 1284).
- e) `cAtivSN` com valor fora de 01, 02, 03, 04 e 90.
Se a DLL valida alguma delas, diga **qual codigo de regra** ela usa (ex.: W59h-10, I17d-xx).

### 4. Ordem das chamadas do emitente
Confirme a ordem: `GerarEmitente`, depois `GerarEmitenteRegimeIBSCBS(1..4)`. Tem que ser nessa ordem? O que acontece se chamar `GerarEmitenteRegimeIBSCBS` sem aplicar (CRT 3 normal)? E se **nao** chamar quando o emitente e do Simples sem regime regular, a DLL acusa erro ou so omite a tag C21a?

### 5. Compra governamental (`tpEnteGov`)
A NT manda trocar `pIBSSN` e `pCBSSN` quando `tpEnteGov` = 1 (Uniao) em 2027 e 2028. A DLL faz essa troca sozinha ao ver `tpEnteGov`, ou o sistema precisa passar os valores ja trocados?

### 6. `cAtivSN` para servico
Em NF-e modelo 55 que traz item de servico (junto com mercadorias), qual codigo de `GerarItensAtividadeSimplesNacional` deve ser usado? A NT so lista 01 a 04 (bens) e 90. Pode-se usar `90`? A DLL aceita?

### 7. CST 620 (combustivel monofasico) no Simples
Para item com CST 620 de emitente do Simples, o sistema nao envia `gTribSN` e envia so `gIBSCBSMono`. A DLL aceita esse item sem `GerarItensImpostoSimplesNacional`? Ela gera o total W59h mesmo assim se nenhum outro item tiver `gTribSN`?

### 8. Valores de `informarIBS` e `cst/cClassTrib` no `GerarItensImpostoSimplesNacional`
No metodo `GerarItensImpostoSimplesNacional(cst, cClassTrib, informarIBS, ...)`, o parametro `informarIBS` deve ser verdadeiro quando? (Pela NT: CRT 1 sem regime regular e CRT 3 pendente pelo CRT 1 informam IBS e CBS; CRT 2 e CRT 3 pendente pelo CRT 2 informam so CBS.) E qual `cst` e `cClassTrib` o sistema deve passar para o Simples, o mesmo que ja usa hoje em `GerarItensImpostoIBSCBS`?

## Formato de resposta desejado

Para cada pergunta (1 a 8):
```
N. SIM / NAO / PARCIAL / NAO SEI
Explicacao curta (ate 3 linhas)
Onde esta no codigo: <classe.metodo ou regra>
```
Se a resposta mudar o jeito de usar a DLL (ex.: parametro obrigatorio, ordem), mostre um **exemplo de chamada em VB6** com os parametros.
