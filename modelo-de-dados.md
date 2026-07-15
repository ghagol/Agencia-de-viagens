# Arquitetura

## Visão geral

Dois ficheiros Google Sheets, ligados num só sentido, mais uma camada de automação em
Apps Script.

```
┌──────────────────────────────────┐        ┌──────────────────────────────────┐
│  📗 OPERACIONAL (toda a equipa)  │        │  🔒 FINANCEIRO (só a gestora)    │
│                                  │        │                                  │
│  Entrada:  Formulário + menu     │        │  Feed:  Importar Reservas        │
│  Dados:    Reservas              │        │         Importar Registos        │
│            Serviços              │──────► │  Cálculo: Comissões              │
│            Passageiros           │ IMPORT │           Custos Fixos           │
│            Registos              │ RANGE  │           Saídas de Caixa        │
│  Catálogos: Clientes, Pacotes,   │(só ler)│  Saída:  P&L · Contas a Receber  │
│             Fornecedores         │        │          Dashboard               │
│  Vistas:   Ficha · Calendário    │        │                                  │
│            CRM · Comprovativo    │        │                                  │
└──────────────────────────────────┘        └──────────────────────────────────┘
              │
              └── Apps Script ──► PDF ──► Drive: «Let's Go - Comprovativos»
```

## Camadas (dentro do Operacional)

| Camada | Abas | Regra |
|---|---|---|
| **Configuração** | `Configurações` | Listas e regras. Mudar aqui muda o sistema. |
| **Catálogos** | `Clientes`, `Pacotes`, `Fornecedores` | Dados mestre, reutilizados. |
| **Transações** | `Reservas`, `Serviços`, `Passageiros`, `Registos` | Dados brutos, append-only. |
| **Vistas** | `Ficha de Reserva`, `Calendário`, `Gestão de Clientes`, `Pagamentos`, `Comprovativo` | Só leem. Reagregam via `FILTER`/`INDEX`. |
| **Entrada** | `Formulário` + menu Apps Script | Ponto único de escrita assistida. |
| **Auditoria** | `Histórico` | Append-only. Escrito pelo script e pelo `onEdit`. |

## Convenção visual

Uma regra que a equipa aprende em 10 segundos:

- **Coluna branca** → escrever.
- **Coluna azul-clara** → calculada, não tocar.

## Fluxo de uma reserva

```
Atendente                          Gestora                      Sistema
    │                                 │                            │
    ├─ Formulário ▸ Guardar reserva ──┼──────────────────────────► Reservas (Aguarda aprovação)
    │                                 │                            └─► Passageiros (titular auto)
    ├─ Adicionar serviço (xN) ────────┼──────────────────────────► Serviços
    │                                 │                            └─► Custo total (SUMIF)
    ├─ Adicionar passageiro (xN) ─────┼──────────────────────────► Passageiros
    │                                 │                            └─► Idade + alerta doc.
    │                                 ├─ Aprovar / definir venda ─► Valor venda + Margem
    │                                 │                            └─► Estado: Aprovado
    ├─ Registar pagamento (xN) ───────┼──────────────────────────► Registos
    │                                 │                            └─► Pago / Pendente / % (SUMIF)
    ├─ Gerar PDF ─────────────────────┼──────────────────────────► Drive (pasta partilhada)
    │                                 │                            │
    │                                 └── Financeiro ◄─ IMPORTRANGE ┘
```

## Segurança e confidencialidade

1. **O Financeiro nunca é partilhado.** É a única garantia real (ver ADR-002).
2. **`IMPORTRANGE` é unidirecional.** Ler não dá acesso à origem.
3. **Colunas calculadas protegidas** em `Dados ▸ Proteger intervalos`.
4. **Limite conhecido e assumido:** o Operacional guarda custo e venda, porque a reserva
   precisa deles para calcular. Confidencialidade total dessas colunas exigiria outra separação.

## Escala

O limite do Google Sheets é de **10 milhões de células por ficheiro** (todas as abas somadas),
não de linhas. Com ~32 colunas, dá centenas de milhares de reservas em teoria; na prática o
desempenho degrada-se por volta das ~100 mil linhas. Para esta agência, são décadas de folga.

Estratégia quando chegar lá: **arquivo por ano** (um ficheiro por ano, o Financeiro lê o
corrente), facilitada pelo modelo append-only. A alternativa de longo prazo é migrar os dados
para uma base de dados e manter o Sheets como camada de relatório.
