# Atualização automática dos clientes — conversa completa, decisões e plano

Documento de passagem entre máquinas. Reúne tudo o que foi conversado e testado no notebook (VB6) em 2026-11-11 sobre fazer o sistema se atualizar sozinho. A sessão que ler isto na máquina com .NET deve tratar este arquivo como a especificação inicial: o que está em "Decisões" já foi acordado com o usuário; o que está em "Em aberto" ainda precisa de resposta dele.

Quando terminar na máquina nova, o usuário volta ao notebook só para **testar** (compilar o VB6 e rodar o piloto). Por isso a Parte 8 lista o que cada máquina faz e a Parte 10 traz o roteiro de testes.

---

## 1. Objetivo

Hoje atualizar ~200 empresas é manual e cansativo: conectar na máquina do cliente, copiar arquivos, instalar a DLL da NF-e, executar scripts do manifesto, trocar os exes. O usuário quer o comportamento de outros programas: ao abrir o sistema aparece "existe uma nova versão", o usuário clica em Sim e o sistema se atualiza sozinho.

## 2. Situação atual (como é feito hoje, à mão)

- Projetos VB6: OnlineCommerce (v3.2.16), PDV/OnlinePDV (v3.3.16), OrdemServico (v1.0.0) e Financeiro (banco próprio `cyber_baseFINANCEIRO`). Cada PC do cliente tem o **seu próprio exe** (não fica em servidor compartilhado).
- Clientes: a maioria tem 1 ou 2 terminais, uns 3 em média, poucos com mais de 4. **Poucos** ainda usam Windows 7 (alguns com SQL Server Express antigo).
- Banco: SQL Server (2008 R2 Express até 2014), acesso por ADO/`oc.ini`, classe `Database.cls`. Um banco por empresa, no PC "servidor".
- Scripts SQL: pasta `Arquivos\scripts` (fonte da verdade, no Git) + `_manifesto.txt` (linhas `número|arquivo.sql|GERAL`; hoje até a 172) + `Executar_Scripts.bat`. São **idempotentes**. Há um espelho byte a byte em `C:\scripts` (a pasta que o usuário leva ao cliente). Regra do repositório: sempre manter `C:\scripts` igual (menos `execucao_*.log`, subpastas `OS`/`SPED`/`Testes`, `Diag_*`, `_AUDITORIA_*`).
- DLL da NF-e/NFC-e: um instalador único (Inno Setup) que popula a pasta `NFE\` dentro da pasta do sistema. Instalado à mão em cada máquina.
- Já existe um padrão de atualização pela nuvem em produção neste projeto: as **alíquotas CBS/IBS** (rotina `AtualizarAliquotasDaNuvem` em `Compartilhado\Modulos\modNFe.bas`): arquivo gerado na aba 3 do `Sistema_Financeiro`, enviado ao Shared Drive, baixado pelo OnlineCommerce/PDV na abertura (1 checagem por dia, controle por `Empresa.AliquotasReformaVersao`/`AliquotasReformaChecagem`, log em `AliquotasReforma_log.txt`, menu manual). **Reaproveitar esse padrão** (Drive, checagem diária, log, menu manual).

## 3. A conversa, em ordem

1. **Pergunta do usuário:** o que precisaríamos fazer para o sistema se atualizar sozinho com aviso e clique em Sim?
2. **Primeira resposta (resumo):** pacote versionado + manifesto no Drive; verificação na abertura; um `Atualizador.exe` separado (o sistema não troca o próprio exe em execução); UAC para registrar DLL; tabela de controle de scripts aplicados; versão mínima do exe no banco; cuidados com TLS 1.2 no Win7, credencial do Google, antivírus, hash, implantação gradual.
3. **Respostas do usuário:** o exe fica em cada PC; ~3 terminais por empresa (maioria 1 ou 2, poucos com mais de 4); poucos clientes em Windows 7.
4. **Ideia do usuário (adotada):** detectar o **servidor**, propor a atualização nele; só o servidor altera o banco; depois que o servidor atualiza, os **terminais são forçados** a se atualizar. A atualização dos terminais é só: exe, DLL da NF-e/NFC-e e arquivos que mudarem. Avaliação: melhor que a ideia inicial (um único ponto para scripts, terminais nunca mexem no banco, coordena a janela entre terminais pela versão mínima gravada no banco).
5. **Prazo (ideia do usuário, adotada):** deve existir **data limite**. Exemplo: algo entra em vigor em 01/12; o sistema avisa desde hoje, o cliente ignora; **1 dia antes** da data limite não há mais opção — aparece a mensagem de que vai atualizar, só o botão OK, e começa o atualizador com a **barra de progresso dele**.
6. **Contexto de rollout:** o usuário vai precisar atualizar todos os clientes agora com tudo que foi criado no ano (todos estão na versão antiga) e quer já levar o controle de versão embutido, para a próxima atualização ser automática. O prazo será definido **no painel** ao criar a atualização.
7. **DLL da NF-e:** é um arquivo único que ele instala manualmente e que popula a pasta `NFE`. Versões em campo: a maioria na **4.4.7.0**, ~10 clientes na **4.5.10.0**, alvo **4.6.0.0**. Atualizar de 4.4.7.0 direto para 4.6.0.0 **não tem problema** (confirmado pelo usuário).
8. **Investigação do instalador:** verificamos metadados e assinatura do arquivo (sem executar) — é **Inno Setup**. O usuário rodou o modo silencioso numa máquina de teste (ver Parte 5).
9. **Sugestão de dividir o trabalho:** o `Atualizador.exe` seria feito na máquina com .NET. O usuário preferiu **passar toda a conversa e o plano** para a outra máquina fazer **tudo**, e depois voltar ao notebook só para testar.

## 4. Decisões já tomadas (pelo usuário)

1. **Servidor primeiro, terminais depois.**
   - Servidor = quem tem o SQL Server rodando local. Só ele roda scripts SQL, e ao final grava no banco a versão nova e a versão mínima aceita do exe.
   - Terminal: na abertura compara a versão do seu exe com a do banco; se estiver defasado, **atualização obrigatória** (sem "Depois"), baixa só exe + DLLs + arquivos locais, sem scripts.
2. **Três estágios, controlados por datas do manifesto** (o exe não tem datas fixas):
   - Opcional: "Atualizar agora?" Sim/Depois, no máximo 1 vez por dia.
   - Insistente: período final (ex.: 7 dias), a cada abertura, com contagem de dias.
   - **Obrigatório: a partir de 1 dia antes da data limite** — mensagem com só o botão OK e o atualizador com barra de progresso começa; não dá para usar a versão antiga.
3. **O prazo é definido no painel** na criação da atualização (versão, motivo, data de aviso, data obrigatória).
4. **Primeira rodada é manual** e já leva a base do controle (tabela de controle, versão mínima, `Atualizador.exe`, credencial de leitura do pacote). A partir daí, tudo pelo painel.
5. **Componente da NF-e via instalador Inno em modo silencioso**, com checagem pela versão do `snfe.dll`.

## 5. O que foi testado e os resultados

### 5.1 Instalador da NF-e (arquivo `Instalador.exe`, 16,2 MB, não assinado digitalmente)

- Metadados: ProductName/FileDescription "OnLine Info EkklesiaSoft Componentes", Company "EkklesiaSoft Desenvolvimento de Sistemas", FileVersion = ProductVersion = **4.6.0.0**. Instaladores antigos têm 4.5.x e menores.
- **Tipo: Inno Setup versão 6.5.2.** (O arquivo do instalador não está no repositório e não deve ser commitado; o usuário fornece.)
- Teste feito pelo usuário numa máquina dele, como administrador:

```
Instalador.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG="C:\tmp\inst.log"
```

- Resultado do log: instalou **sem nenhuma janela**, em ~8 segundos, `Installation process succeeded`, código de saída **0**, `Need to restart Windows? No`.
  - Roda com `User privileges: Administrative` / `Install mode root key: HKEY_LOCAL_MACHINE` → **precisa de UAC**.
  - Usou o **diretório da instalação anterior** (ex.: `C:\Projeto\OnlineCommerce\NFE`, de onde já havia `unins000.dat`). Num cliente sem instalação anterior igual o destino pode ser outro → **sempre passar `/DIR=`**.
  - **Sobrescreve todos os seus arquivos sempre** (mesmo com timestamp igual); não é incremental. A lista é só dos arquivos do próprio instalador: `snfe.dll`, `snfe.danfe.dll`, `snfe.gnre.dll`, `zapzap.dll`, `GoogleDriveUploader.dll`, `Newtonsoft.Json.dll`, `QRCoder.dll`, `FastReport*.dll`, DLLs System.*/Microsoft.* do .NET, `RegAsm.exe`, `registraDLL*.bat`/`desregistraDLL*.bat`, `NFE\nfe\schemas\*`, cadeia de certificados em `NFE\nfe\cadeiascertificado\`, fontes `.ttf` (em `NFE\fontes` e também em **`C:\Windows\Fonts`**).
  - **Não toca** em arquivos do cliente: nada de `snfe.config`, certificado do cliente, XMLs, `ConsProt`, `Backup`, `ExportarXML`. Isso afasta o risco já visto de corrupção do `snfe.config`.
  - Ao final **registra sozinho as DLLs** com `C:\Windows\Microsoft.NET\Framework\v4.0.30319\Regasm.exe <dll> /tlb:<dll>.tlb /codebase` para `snfe.dll`, `zapzap.dll` e `GoogleDriveUploader.dll`, todos com código de saída 0. O atualizador **não precisa** de `regsvr32` para essas três.
  - Usa o Restart Manager do Windows ("46 files to register"); se algo estiver usando as DLLs, ele tenta lidar, mas o atualizador deve **fechar o sistema antes**.
  - Cria a chave de desinstalação `HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\OnLine Info EkklesiaSoft Componentes_is1`. Como o instalador é de 32 bits, **a leitura precisa olhar `WOW6432Node`** (`DisplayVersion` = 4.6.0.0). A chave sem `WOW6432Node` não existe.
- **Fonte de versão confiável:** `FileVersion` de `<pasta do sistema>\NFE\snfe.dll` (=4.6.0.0 depois da instalação). **`snfe.gnre.dll` tem versão própria (1.5.0.0)** — não esperar 4.6.0.0 nele.
- Efeito colateral aprendido: rodar o instalador sobre a pasta do repositório altera 13 arquivos versionados (`snfe.dll`, `snfe.gnre.dll`, `snfe.tlb`, 8 DLLs .NET, `unins000.*`). Para testar sem sujar o repositório, usar `/DIR="C:\Temp\NFE_teste"`.
- Foi validado o funcionamento da 4.6.0.0 pelo usuário (emissão), e os arquivos já foram commitados (commit `6390452`, `unins000.*` de fora). `NFe_Completa`/`modNFe`/`Notas_Destinadas_Consulta` passaram a imprimir o DANFE com `modelo = 4` (layout novo) — commit `6a446dd`.

### 5.2 O que NÃO foi testado

- Nada do atualizador em si (não existe ainda).
- Instalação silenciosa sobre uma versão antiga (4.4.7.0) para ver se o resultado é igual (o usuário disse que a atualização direta não tem problema; falta o teste automatizado).
- Instalação silenciosa com `/DIR=` explícito em pasta nova.
- Comportamento com o sistema aberto (DLL travada).
- Win7 (TLS 1.2, .NET).

## 6. Dados técnicos que a outra sessão vai precisar

### 6.1 Google Drive (já em uso)

- Shared Drive **"IBPT"**, ID `0AI0VAvFMSupDUk9PVA`. Pastas existentes: IBPT (tabela do IBPT) e ALIQUOTAS (criada à mão; o arquivo de alíquotas é lido pelos clientes). Dono do drive: `backup@online-sistemas.com`.
- Componente COM `GoogleDriveUploader.Uploader` (`GoogleDriveUploader.dll`, .NET, usa `Google.Apis.Drive.v3`): métodos `UploadFileToFolderName`, `DownloadLatestFileByPath` (baixa o arquivo **mais recente** da pasta). Em VB6: `GoogleEnviarArquivo` / `GoogleBaixarArquivo` em `Compartilhado\Modulos\Util.bas` (o PDV tem `Util.bas` próprio, sem `GoogleBaixarArquivo`; o download foi implementado dentro do `modNFe.bas`).
- Credencial atual dos clientes: arquivo `.json` de uma conta de serviço (`App.Path\NFE\backupclientes-488021-067a35f7c315.json`). **Essa conta tem escrita no Drive** (usada pelo backup em nuvem). **Não colocar chave privada neste repositório.**
- **Recomendação (a implementar):** criar **outra conta de serviço, somente leitura**, em uma pasta separada `ATUALIZACOES`, e distribuir só essa credencial no pacote do atualizador. Assim um cliente comprometido não consegue trocar o pacote dos outros 199.

### 6.2 Pontos do código VB6 a ler

- `EhServidorLocal()` (usada em `OnlineCommerce\Forms\Tela_Principal.frm`, `MDIForm_Load`): decide hoje o que é "servidor". **Ler antes de usar** — a regra segura proposta é "quem tem o SQL Server rodando local".
- `AtualizarAliquotasDaNuvem` / `AliqNuvem_*` em `Compartilhado\Modulos\modNFe.bas`: modelo da checagem diária, log e menu manual.
- `OnlineCommerce\Forms\Tela_Principal.frm` e `PDV\Forms\PDV.frm` (fim do `Form_Load`): onde a checagem de versão entraria na abertura.
- `Financeiro\Projeto\Sistema_Financeiro.frm`, aba 3 "ALÍQUOTAS CBS/IBS": modelo de painel que monta arquivo e envia ao Drive.
- Tabela `configuracao` (`config_nome`, `config_valor`): padrão para guardar parâmetros (ex.: `SPED_ENVIO`, `SPED_MENSAL`). Usar para `VERSAO_SISTEMA` e `VERSAO_MINIMA_EXE`.
- `Executar_Scripts.bat` e `_manifesto.txt` em `Arquivos\scripts`: **ler** como os scripts são executados hoje (credencial do SQL, `sqlcmd`, ordem, tratamento de erro) antes de reproduzir isso no atualizador.

### 6.3 Regras do repositório (CLAUDE.md) para qualquer código VB6

- `.frm/.bas/.cls` são cp1252: editar **sempre por Python em modo binário** (`rb`/`wb`), normalizar CRLF e conferir `count == 1` antes de substituir; **nunca** usar Edit/Write direto neles se tiverem acento.
- Limite de 254 controles por form; `IIf` avalia os dois lados; `Const`/variáveis de módulo não podem ficar entre Subs; transações com `dbData.Execute "BEGIN TRANSACTION"` e ROLLBACK no handler de erro.
- **Não fazer commit/push sem pedido explícito do usuário.** Não commitar `PDV/Config.ini`.
- Scripts SQL novos: entrar no `_manifesto.txt` e ser espelhados byte a byte em `C:\scripts`.
- O usuário exige **respostas em português**.

## 7. Proposta de contrato entre as partes (a validar/ajustar)

### 7.1 Manifesto (JSON publicado no Drive)

```json
{
  "versao": "3.4.0",
  "motivo": "Vigência da CBS em 01/12",
  "aviso_a_partir_de": "2026-11-20",
  "obrigatoria_em": "2026-12-01",
  "pacote": "Atualizacao_3.4.0.zip",
  "arquivos": [
    {"destino": "OnlineCommerce.exe", "sha256": "…", "tipo": "exe", "alvo": "todos"},
    {"destino": "PDV.exe", "sha256": "…", "tipo": "exe", "alvo": "todos"},
    {"destino": "NFE\\Instalador_4.6.0.0.exe", "sha256": "…", "tipo": "instalador-nfe", "versao": "4.6.0.0"}
  ],
  "scripts": {"ate_numero": 172, "pasta": "scripts"},
  "versao_minima_exe": "3.4.0"
}
```

Regras: o cálculo do estágio (opcional / insistente / obrigatório) é feito pelo cliente a partir de `aviso_a_partir_de` e `obrigatoria_em`; a **data de referência vem do SQL Server** (`GETDATE()` do servidor) ou do servidor do Google — **nunca do relógio do PC**, para ninguém escapar do prazo atrasando o relógio. Guardar a última data obrigatória conhecida para o bloqueio valer mesmo sem internet.

### 7.2 Banco

- Tabela `ScriptsAplicados` (`Numero INT PK`, `Arquivo`, `DataAplicacao`, `Resultado`, `Maquina`). O atualizador só roda os scripts com número maior que o último aplicado. **Marco zero:** na primeira rodada, rodar todos os scripts do manifesto (idempotentes) e registrar todos como aplicados.
- Em `configuracao`: `VERSAO_SISTEMA` (gravada pelo servidor ao terminar) e `VERSAO_MINIMA_EXE` (o exe se recusa a abrir abaixo disso e oferece/força a atualização).
- Scripts continuam **apenas aditivos** (não remover nem renomear colunas) para o exe antigo conviver com o banco novo durante a janela entre terminais.

### 7.3 `Atualizador.exe`

- Recomendação: **.NET Framework 4.8 (ou 4.7.2), um único exe, sem instalação**, com manifesto de aplicativo `requireAdministrator` (UAC). **Não usar .NET moderno (6/8):** não roda em Windows 7. Os clientes já têm o CLR 4.0.30319 (a DLL da NF-e depende dele). Precisa de TLS 1.2 habilitado para falar com o Google (verificar no Win7).
- Chamado pelo sistema VB6 por `Shell`, com parâmetros do tipo `/modo=servidor|terminal /pasta="<pasta do sistema>" /servidor=<instância SQL> /manifesto=<url ou caminho>`. O sistema VB6 fecha em seguida; o atualizador espera o processo terminar, faz o trabalho e reabre o sistema.
- Passos (modo servidor): 1) baixar o pacote e conferir **SHA-256** de cada arquivo; 2) backup do exe atual (e do banco, se couber no ambiente do cliente); 3) rodar o instalador da NF-e se `snfe.dll` < versão do manifesto (`/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /DIR="<pasta>\NFE" /LOG=...`, esperar código de saída 0 e reler a versão do `snfe.dll`); 4) rodar scripts pendentes e registrar em `ScriptsAplicados`; 5) trocar exes; 6) gravar `VERSAO_SISTEMA`/`VERSAO_MINIMA_EXE` **só se tudo deu certo**; 7) reabrir o sistema.
- Modo terminal: passos 1, 2, 3 (se a DLL mudou), 5 e 7; não toca no banco.
- **Rollback:** se qualquer passo falhar, restaurar o exe do backup, **não gravar a versão nova** (os terminais não podem tentar atualizar para uma versão quebrada), registrar em log e mostrar mensagem clara. Falha da atualização nunca deve impedir o uso da versão anterior.
- Códigos de saída (proposta): 0 sucesso; 10 sem internet; 11 hash inválido; 12 instalador da NF-e falhou; 13 script SQL falhou; 14 troca de exe falhou; 20 cancelado/sem permissão.
- Log: arquivo por execução (ex.: `Atualizador_AAAAMMDD_HHMMSS.log`) na pasta do sistema, no estilo dos outros `*_log.txt`.
- Barra de progresso e mensagens em português.

### 7.4 Painel de publicação

- Aba "Atualizações" no `Sistema_Financeiro` **ou** uma ferramenta .NET separada ("Publicador"): escolhe os arquivos do pacote, calcula o SHA-256, monta o `manifesto.json` e o zip, define versão, motivo, `aviso_a_partir_de` e `obrigatoria_em`, e envia para a pasta `ATUALIZACOES` do Drive. Calcular hash em VB6 é trabalhoso; a recomendação é a ferramenta .NET (a aba do Financeiro pode só chamá-la).
- Acompanhamento (fase posterior): terminais gravam a versão em que estão, para o usuário ver quem já atualizou e ligar para quem ficou para trás antes do prazo.

## 8. Divisão do trabalho

**Máquina com .NET (esta sessão faz tudo o que for possível sem o IDE do VB6):**
- `Atualizador.exe` (modos servidor e terminal, UAC, barra de progresso, hash, rollback, log).
- Ferramenta de publicação (manifesto + zip + upload).
- Quando o código VB6 for só texto, pode escrever os trechos novos seguindo as regras da Parte 6.3: tabela `ScriptsAplicados` + script SQL (com entrada no manifesto e espelho em `C:\scripts`), checagem de versão e dos 3 estágios na abertura do OnlineCommerce e do PDV, chamada do atualizador, menu manual "Verificar atualização".
- **Não compilar o VB6** — isso é no notebook.

**Notebook (VB6) na volta:**
- Compilar OnlineCommerce, PDV, OrdemServico e Financeiro e corrigir erros de compilação.
- Rodar o piloto (Parte 10).

## 9. Em aberto (decidir com o usuário)

1. **Linguagem do `Atualizador.exe`:** recomendação .NET Framework 4.8 (acima). O usuário ainda não confirmou formalmente; a ideia de usar a máquina .NET indica que sim.
2. **Como detectar o servidor** na prática (ler `EhServidorLocal()`; casos de 1 PC que é servidor e terminal, e de SQL em máquina que não é terminal).
3. **Credencial de leitura** para o pacote (criar a segunda conta de serviço e a pasta `ATUALIZACOES`).
4. **Backup do banco antes dos scripts:** o ambiente dos clientes (SQL Express) permite? Onde guardar? Alternativa mínima: scripts aditivos + parar na primeira falha + não gravar versão.
5. **Assinatura digital** do `Atualizador.exe` (e o instalador da NF-e não é assinado) para evitar alertas do SmartScreen/antivírus.
6. **Janela de implantação:** piloto com 2 ou 3 clientes de confiança (incluindo um Win7) antes dos 200.
7. **Horário do estágio obrigatório:** preferir forçar na primeira abertura do dia, não no meio do pico de vendas; nunca interromper uma venda do PDV.
8. **Pendências da versão que vai como "marco zero":** antes de montar o primeiro pacote, estabilizar e testar o que foi criado no ano (recompilar tudo; testes de campo ainda pendentes: `/enviarxml` compilado, `.bat` com UAC, Pix dinâmico, fechamento de OS com 3 formas, SPED e alíquotas CBS/IBS pela nuvem, `Configuracao_Geral` com o `chkEnviarSpedMensal` na aba GERAL).

## 10. Roteiro de testes (na volta ao notebook)

1. **Atualizador isolado:** rodar `Atualizador.exe` em modo servidor numa cópia local, com `/DIR` de teste, e conferir: hash, instalação da NF-e (versão do `snfe.dll`), scripts aplicados em `ScriptsAplicados`, troca do exe, versão gravada no banco.
2. **Rollback:** simular falha (hash errado, script com erro, instalador com código de saída ≠ 0) e conferir que o exe antigo volta e a versão **não** é gravada.
3. **Terminal:** com o banco já na versão nova, abrir um exe antigo e conferir que a atualização é obrigatória e não mexe no banco.
4. **Estágios:** simular datas (parâmetro de teste que avança o "hoje") e conferir opcional → insistente → obrigatório (só OK + barra).
5. **Sem internet** e com relógio do Windows adiantado/atrasado.
6. **NF-e:** depois de atualizar, emitir NF-e e NFC-e em homologação e imprimir o DANFE (modelo 4).
7. **Win7:** testar TLS 1.2 e o `Atualizador.exe` em um Windows 7.
8. **Piloto real:** 2–3 clientes de confiança, com acompanhamento, antes dos demais.

## 11. Resumo executivo

- Servidor atualiza primeiro (banco só nele), terminais depois, obrigatórios.
- Três estágios por data do manifesto; no último dia só OK e barra de progresso.
- Instalador Inno da NF-e roda em silêncio, registra as próprias DLLs e não mexe na configuração do cliente; versão pelo `snfe.dll`.
- Atualizador em .NET Framework 4.8 (Win7-compatível), UAC, hash, rollback, log.
- Painel define versão, motivo e prazos; publica no Drive em pasta com credencial só de leitura.
- Primeira rodada manual já leva a base do controle; depois tudo automático.
