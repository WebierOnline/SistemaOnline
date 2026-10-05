# Plano: Módulo Mesas/Comandas + Migração futura pra C#/.NET

> Documento de planejamento — nada foi implementado ainda. Escrito em 2026-09-26 pra dar
> contexto a uma sessão nova do Claude Code (ex: rodando no notebook Samsung reservado pra
> esse projeto), sem precisar o usuário reexplicar tudo do zero.

## Contexto do negócio

O usuário é dono/desenvolvedor solo de um sistema comercial (ERP/PDV) escrito 100% em VB6 +
SQL Server Express, vendido/instalado pra vários clientes reais (cada cliente com seu próprio
servidor SQL Server Express local — não é um banco central compartilhado). Projetos
existentes no mesmo repositório Git (`C:\Projeto`, remoto
`https://github.com/WebierOnline/SistemaOnline.git`):

- **OnlineCommerce** — sistema principal (cadastro, NFe/NFCe, financeiro, etc.)
- **PDV** — ponto de venda (roda no mesmo processo/instância que várias telas do
  OnlineCommerce em alguns clientes)
- **OrdemServico** — módulo de OS (automotivo/recapadora, entre outros tipos de empresa)
- Módulo **Compartilhado** (Forms/Classes/Módulos reusados entre os projetos acima)

Instalou recentemente um cliente novo: um **restaurante**, vendendo/usando o PDV do jeito que
está hoje (venda tipo balcão de supermercado). Percebeu que faltaria um controle de
**mesas/comandas** — recurso comum em sistemas de bar/restaurante que o PDV atual não tem.

## A ideia nova: módulo de Mesas/Comandas

Usuário viu um vídeo de um sistema concorrente (ST3 Sistemas) com uma tela de "mapa de
mesas": grade numerada de mesas, cor por status (livre/ocupada/reservada), mostrando qtd de
pessoas, tempo decorrido e valor acumulado por mesa, com teclado numérico e ícones de ação
(imprimir conta, cozinha, fechar, etc.) — imagem de referência `image287.jpg` em
`Arquivos/`.

**Lista de recursos extraída** (2026-09-28, a partir da transcrição de um segundo vídeo do
mesmo fornecedor — `Arquivos/tactiq-free-transcript-ZNTWpdMEA5c.txt`), organizada por grupo
funcional:

- **Modos de operação**: controle por mesa (mapa do salão) OU por comanda avulsa — os dois
  coexistem no mesmo sistema. Uma mesma mesa pode ter várias comandas, ou ser dividida por
  nome de cliente dentro da mesma mesa.
- **Reserva de mesa**: selecionar cliente existente ou cadastrar um novo na hora; reservar
  uma mesa específica pra ele; mesa muda de status/cor pra "reservada".
- **Agrupar mesas** (ex: mesa 6 + mesa 7): lançamentos em qualquer uma das duas mesas
  vinculadas caem na mesma conta.
- **Transferir itens ou a mesa inteira** de uma mesa pra outra (cliente mudou de lugar) —
  suporta transferência parcial (só alguns itens) ou total.
- **Fechamento pra conferência** (antes do pagamento de fato): imprime um resumo com os
  itens por cliente, já somando taxa de serviço e "cover artístico" quando aplicável; a mesa
  muda de cor/status pra "fechada" nesse momento (ainda não é o pagamento).
- **Recebimento/pagamento**:
  - Mesa dividida por cliente: selecionar o nome carrega só os itens daquela pessoa pro
    lado do recebimento.
  - Pagamento parcial de item compartilhado (ex: dividir o valor de 1 porção entre 2
    pessoas).
  - Agrupar várias comandas separadas no momento do pagamento (ex: família com 3 comandas
    querendo pagar tudo junto).
  - Pagamento misto (parte dinheiro + parte cartão) na mesma finalização.
  - Emite recibo/cupom fiscal ao finalizar.
- **Escopo maior do fornecedor** (mencionado no fim do vídeo, não exclusivo do módulo de
  mesas): delivery, venda balcão, estoque, financeiro, fiscal — reforça a decisão já tomada
  de que o módulo novo **não deve duplicar** isso; delivery/balcão/estoque/financeiro/fiscal
  já existem no OnlineCommerce/PDV, só precisa a comanda fechada desaguar nesse fluxo
  existente (ver "Ponto de integração crítico" acima).

**Cruzamento com o checklist de um site sobre softwares de restaurante** (2026-09-28,
verificado item a item no código atual, não é achismo):

- **Já existe no VB6, só precisa integrar (não redesenhar)**: PDV com turno (abrir/fechar
  caixa, fundo de troco, sangria, suprimento — `Caixa_Fechamento.frm`/`Caixa_Retirada.frm`/
  `Caixa_Suprimento.frm`/`Caixa_Controle.frm`), troco na tela (`txtTroco` no PDV), NFC-e em
  vários modos (maduro, muito trabalhado nesta sessão), estoque baixa na venda + alerta de
  mínimo (`Estoque_Minimo.frm`).
- **Gap real, não existe, precisa nascer no projeto novo**: comanda/mesa + app do garçom (o
  próprio projeto), KDS + ticket de alteração/cancelamento de item, cardápio digital que
  aceita pedido (sem cadastro/app), delivery com taxa por região + acerto de entregadores,
  **ficha técnica e CMV** (relação prato→insumos+quantidade — o gap mais importante da lista,
  não é exclusivo de mesas/comandas, é buraco de qualquer prato pronto vendido), fiado/
  mensalista (existe crediário "À Prazo" parcial, mas não "conta corrente do cliente" tipo
  marmitaria), maquininha integrada (TEF).
- **Parcial/não confirmado**: pagamento dividido numa mesma finalização (dinheiro+Pix+cartão
  simultâneos) — não confirmado se o balcão já faz isso hoje; notas do mês pra contabilidade
  em CSV/Sintegra — existe exportação de XML de notas, formato Sintegra especificamente não
  confirmado.
- **Corte de escopo proposto** (usuário ainda não confirmou): entra no desenho agora =
  comanda/mesa + app garçom + KDS básico + cardápio digital + pagamento dividido. Fica pra
  depois mas o schema já reserva espaço = ficha técnica/CMV, delivery+entregadores, fiado/
  mensalista, TEF. Fora de prioridade pro público do usuário (pizzaria/lanchonete/padaria) =
  recursos mais de bar/balada (ex: rodízio, comanda por pulseira/RFID).

**Os 3 canais de venda** (ponto levantado pelo usuário, 2026-09-28): toda venda cai em 1 de 3
canais — **mesa** (local, ocupa mesa/comanda), **balcão** (local, sem ocupar mesa — é o fluxo
atual do PDV, inalterado), **delivery** (internet/remoto). Os três devem convergir pro mesmo
`pedidos`/`pedidos_itens` no fechamento (reaproveitando caixa/NFCe/financeiro existente), mas
tem "vida" diferente antes disso.

**Achado técnico importante pra quando formos desenhar**: `pedidos.tipo_pedido` só é
preenchido no FECHAMENTO da venda (fica vazio/NULL enquanto o pedido está aberto — é esse
vazio que `ExistePedidoLivre` usa pra achar um pedido "livre" pra retomar). Por isso, **não
dá pra marcar a origem mesa/delivery só com um valor novo em `tipo_pedido`** — durante a
fase aberta (que é quando mais importa saber a origem), o campo ainda estaria vazio, igual
balcão. A forma correta é marcar a origem pela **existência do vínculo**: se o pedido tem
uma linha em `Comandas` (FK `cod_pedido`) é venda de mesa; se tem linha em `PedidosDelivery`
(nome provisório) é delivery; se não tem nenhum dos dois, é balcão (como hoje). No
fechamento, os três continuam gravando `tipo_pedido = 'VENDA'` igual, sem ensinar o sistema
antigo a entender canal nenhum. Quando formos implementar de verdade, `ExistePedidoLivre` e o
bloqueio de fechamento de caixa (`Caixa_Fechamento.frm`) vão precisar de mais uma condição de
exclusão (não pegar pedido com Comanda/Delivery vinculado).

Levantamento de recursos e considerações de arquitetura ainda em andamento — usuário optou
por continuar trazendo pontos antes de eu desenhar o schema de verdade. **Não iniciar o
desenho ainda.**

## Decisão de arquitetura (já fechada)

**Não construir esse módulo em VB6.** Motivo: o plano de longo prazo do usuário (ver seção
seguinte) inclui dar acesso mobile/web a esse e outros módulos no futuro — VB6 não tem
caminho nenhum pra mobile/web, então construir em VB6 agora significa reescrever do zero
depois. Como é um módulo novo (sem dívida técnica herdada), é o lugar ideal pra já nascer na
stack nova.

**Arquitetura recomendada** (camadas):

```
SQL Server (mesmo banco/mesma instância que o VB6 já usa — nada muda aqui)
        ↑
   API nova em C#/.NET (ASP.NET Core Web API) — peça que ainda não existe
        ↑ ↑ ↑
  PDV/Caixa (VB6, continua falando direto com o SQL Server como hoje)
  App do garçom (celular — via navegador, Blazor resolve isso sem precisar de app nativo já de cara)
  Painel do dono (web, olhando de casa)
```

- Banco de dados: **continua SQL Server**, sem trocar de tecnologia. É o que o usuário já
  domina profundamente (scripts, deploy, ADO, etc.).
- Linguagem escolhida pra API/novo desenvolvimento: **C# / ASP.NET Core** — integração nativa
  com SQL Server (mesmo fabricante), curva de aprendizado mais suave vindo de VB6 (mesmo
  "mundo" Microsoft/Visual Studio) do que ir direto pra Node/Python/etc.
- **Ponto de integração crítico**: uma comanda fechada precisa virar uma venda de verdade nas
  MESMAS tabelas que o balcão já usa hoje (`pedidos`/`pedidos_itens`), pra reaproveitar todo o
  fluxo de fechamento de caixa, NFCe, relatórios já existente e testado — não criar um
  universo de dados paralelo e desconectado. Isso precisa ser levado em conta no desenho do
  schema de `Mesas`/`Comandas`/`ComandaItens` quando chegar a hora.
- O PDV em VB6 **não precisa mudar nada pra isso funcionar** — continua lendo/escrevendo
  direto no SQL Server como sempre fez. A API nova é só mais uma porta de entrada pro mesmo
  banco, pros clientes que não podem/devem falar direto com o SQL Server (celular, web).

## Plano faseado (evitar tentar tudo de uma vez)

1. **Fase 0 — aquecimento**: programa pequeno e descartável em C# conectando num banco SQL
   Server que o usuário já conhece de cor (ex: cópia local de `cyber_base`), fazendo um CRUD
   simples (listar produtos, cadastrar categoria). Objetivo: se acostumar com sintaxe C# e
   Visual Studio usando dados familiares, não tutorial genérico.
2. **Fase 1 — produto mínimo real**: API (ASP.NET Core) + **uma única interface web**
   (recomendado **Blazor** — permite escrever a UI inteira em C#, sem aprender
   JavaScript/React em paralelo; como é uma página web comum, já dá acesso via navegador do
   celular também, sem precisar construir app nativo ainda). Validar o fluxo completo
   (mesas → comanda → itens → fechar) com o cliente do restaurante.
3. **Fase 2 — produtizar**: adaptar pra vender aos outros clientes de pizzaria/lanchonete/
   padaria que o usuário já tem, e novos que surgirem.
4. **Depois (não urgente)**: app nativo de verdade via **.NET MAUI** (mesma linguagem C#,
   funciona offline, mais "cara de app"), quando a API já estiver madura.

## Plano de longo prazo (mencionado pelo usuário, ainda não detalhado)

O usuário já tinha planos de eventualmente **migrar os outros sistemas** (OnlineCommerce/PDV/
OrdemServico) pra essa mesma stack C#/.NET também — inclusive pra oferecer acesso mobile/web
a algumas funções pros clientes que já usam o sistema VB6 hoje. Esse módulo de Mesas/Comandas
é visto como o **primeiro passo prático** desse plano maior (aprender a stack num projeto
novo, de baixo risco, antes de pensar em tocar no sistema legado de 15 anos em produção).
Nada disso foi detalhado ainda além dessa intenção geral.

## Infraestrutura de desenvolvimento (decidido)

Preocupação do usuário: não quer arriscar quebrar o ambiente VB6/OCX/SQL Server Express que
já funciona bem na máquina de trabalho principal (notebook, Windows 10 Pro, i5-7200U, ~20GB
RAM, 2 SSDs físicos) instalando Visual Studio/.NET nela.

**Decisão**: usar uma máquina **separada e isolada** pro trabalho em C#/.NET — nunca instalar
nada de .NET/Visual Studio na máquina principal. Duas opções discutidas:

- VM via Hyper-V (grátis, nativo no Windows Pro) — solução de custo zero, mas some junto se o
  SSD principal morrer, a menos que o `.vhdx` fique no segundo SSD físico da máquina (o
  usuário tem 2 discos disponíveis).
- **Escolhida**: notebook **Samsung parado, sem uso**, que o usuário já possui — resolve
  isolamento de software E proteção contra falha de hardware ao mesmo tempo (é fisicamente
  outra máquina), sem custo.

**Specs do Samsung** (foto verificada): `DESKTOP-10TK0E9`, Intel Core i3-3110M @ 2.40GHz
(dual-core, ~2012/2013), 4GB RAM instalados (3,88GB utilizável), Windows 10 Pro 64 bits.
Usuário vai verificar se tem HD ou SSD, e pretende colocar SSD + mais 4GB de RAM (totalizando
8GB). Confirmado: é notebook (tela/teclado integrados).

**Avaliação técnica dada**: com SSD + 8GB RAM, o Visual Studio 2022 completo RODA, mas o
processador (dual-core de 2012) continua sendo o gargalo real — vai ser mais lento que a
máquina principal, mas perfeitamente utilizável pra projeto pequeno/médio (não é
"impossível", é "não instantâneo"). Recomendações concretas dadas:
- Na instalação do VS2022, marcar só as cargas de trabalho necessárias (".NET desktop
  development" + "ASP.NET and web development"), não instalar cargas extras (jogos, Azure,
  mobile) que pesam em segundo plano.
- Se sentir pesado no dia a dia, alternativa mais leve: **VS Code + extensão C# Dev Kit** (o
  mesmo código, IDE mais leve) — pode ser a escolha inicial mesmo antes de qualquer
  necessidade, pra aproveitar melhor o hardware mais fraco.

## Sincronização entre as duas máquinas (decidido)

**Git é o mecanismo — nada de ferramenta nova.** O `C:\Projeto` já é um repositório Git com
push pro GitHub. No Samsung, clonar o mesmo jeito (`git clone`) em vez de copiar pasta
manualmente. Disciplina: `git push` ao terminar de mexer numa máquina, `git pull` antes de
começar a mexer na outra.

**Evitar explicitamente**: NÃO usar OneDrive/Google Drive/Dropbox pra sincronizar a pasta do
projeto por cima do Git — risco real de corromper os arquivos binários `.frm`/`.frx`
(sensíveis a codificação cp1252, ver `CLAUDE.md`) ou embolar a pasta `.git`, já que duas
ferramentas de sync brigando pelos mesmos arquivos é receita de problema.

**Estrutura de repositórios recomendada**: o projeto novo (C#/.NET) deve ser um
**repositório Git separado** do `SistemaOnline` (stack diferente, ciclo de vida diferente,
pode virar produto separado no futuro) — não misturar dentro do repo VB6 atual. No Samsung,
faz sentido clonar os dois: o `SistemaOnline` (só como referência de leitura, pra consultar
schema de tabelas como `pedidos`/`produtos` na hora de desenhar a integração) e o repositório
novo (onde o trabalho de verdade acontece).

## Sobre a conta/plano Claude

Confirmado: o plano do Claude é vinculado à CONTA (login), não à máquina — dá pra usar Claude
Code (mesma conta) nas duas máquinas sem custo adicional. O uso (mensagens/limite) vem de um
único "balde" compartilhado por conta, não um balde por dispositivo — usar as duas ao mesmo
tempo intensamente divide o mesmo limite, não dobra a capacidade, mas isso não deve ser
problema real pro uso esperado (uma máquina por vez, na prática).

**Nota pra quem ler isso numa sessão nova**: uma sessão do Claude Code no Samsung começa com
memória própria (não herda automaticamente as memórias/notas acumuladas na sessão do
notebook principal) — mas como o código fica sincronizado via Git, o estado do projeto em si
está sempre atualizado independente disso. Esse arquivo aqui é exatamente a ponte pra não
perder o contexto da CONVERSA (decisões, plano, motivação) entre as duas sessões.

## Próximos passos (nada disso foi feito ainda)

1. ~~Usuário vai reassistir o vídeo de referência e listar os recursos da tela de mesas com
   mais detalhe.~~ **Concluído 2026-09-28** — ver lista de recursos na seção acima.
2. **Próximo passo atual**: desenhar o schema (`Mesas`, `Comandas`, `ComandaItens`, e as
   entidades que a lista de recursos acima exige — reserva, vínculo entre mesas, histórico de
   transferência, split de pagamento por pessoa/item) e como ele se conecta com
   `pedidos`/`pedidos_itens` existentes. Ainda não iniciado.
3. Hardware: usuário está cotando um computador novo (desktop) pra substituir o notebook
   Samsung como máquina dedicada — configuração já avaliada e aprovada (ver seção de
   Infraestrutura): i5-12400, 32GB RAM DDR4, SSD NVMe 500GB, placa-mãe H610, gráficos
   integrados, fonte 500W real. Softwares necessários (Visual Studio/.NET/SQL Server
   Express/Git/Android SDK) confirmados gratuitos pro uso comercial da empresa do usuário;
   únicos custos futuros e opcionais: taxa única de US$25 na Google Play (só se publicar o
   app na loja pública — pode não ser necessário se for só uso interno dos funcionários do
   cliente) e US$99/ano da Apple (só se decidir publicar em iOS).
4. Nenhum código foi escrito ainda — essa conversa inteira foi só de planejamento/decisão de
   arquitetura, infraestrutura e levantamento de requisitos.
