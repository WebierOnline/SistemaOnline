# OnlineCommerce — Sistema VB6

## Estrutura do projeto

A pasta de trabalho do projeto VB6 é `C:\Projeto\` (raiz do repositório).

```
C:\Projeto\
├── Compartilhado\          ← classes, forms e módulos compartilhados entre projetos
│   ├── Classes\
│   ├── Forms\
│   └── Modulos\            (General.bas, Util.bas, JSON.bas, modNFe.bas)
├── OnlineCommerce\         ← OnlineCommerce.vbp  (v3.2.16) — sistema principal
│   ├── Classes\
│   ├── Forms\
│   └── Modulos\
├── PDV\                    ← OnlinePDV.vbp  (v3.3.16) — ponto de venda
│   ├── Classes\
│   ├── Forms\
│   └── Modulos\            (inclui lepeso.bas para balança via MSComm32)
├── OrdemServico\           ← OrdemServico.vbp  (v1.0.0) — OS automotivo/recapadora
│   ├── Classes\
│   ├── Forms\
│   └── Modulos\
└── Financeiro\             ← módulo financeiro (projeto a criar)
    ├── Classes\
    ├── Forms\
    └── Modulos\
```

## Regras críticas para edição de .frm

- Arquivos `.frm` são **Windows-1252 (ANSI)**. Editar SEMPRE via Python em modo binário (`rb`/`wb`).
- Normalizar CRLF após cada edição: `.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")`
- **NUNCA** usar o Edit tool diretamente em `.frm` — corrompe a codificação.
- Sempre verificar unicidade do trecho antes de substituir (`data.count(old) == 1`).

## Banco de dados

- SQL Server 2014
- Conexão via ADO 2.8 — classe `Database.cls`
- Configuração em `oc.ini`
- Todos os projetos compartilham a mesma base de dados
- `dbData As Database` (DAO) — usado para `Execute "BEGIN/COMMIT/ROLLBACK TRANSACTION"`
- `RsOpen` / `SQLExecutaRetorno` — funções utilitárias ADO em `General.bas`

## Padrões de código VB6

- Transações: `dbData.Execute "BEGIN TRANSACTION"` com flag `bTrans As Boolean` + `ErrHandler` com ROLLBACK
- **Deadlock**: nunca abrir `RsOpen` (ADODB) dentro de transação DAO — mover leituras para ANTES do BEGIN
- Sempre capturar `Err.Description` em variável local ANTES de `On Error Resume Next` no handler
- `RecordCount` em DAO requer `MoveLast` — ou usar `SELECT COUNT(*)` via `SQLExecutaRetorno`

## Git

- Não fazer commit/push automaticamente — somente quando o usuário pedir explicitamente.
- Repositório: https://github.com/WebierOnline/SistemaOnline.git
