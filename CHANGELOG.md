# Changelog

## v3.0 — Robustez e auditoria

- `LockService` nas operações que geram IDs (evita duplicados com vários atendentes)
- Rollback quando a criação de uma reserva falha a meio
- Histórico completo: ações do menu **e** edições manuais, com valor antes/depois
- Validações: datas, prestações, custos, e pagamento nunca superior ao pendente
- Nova ação `Atualizar cliente`
- P&L anual dinâmico (ano definido em `Configurações`)
- Caixa calculado a partir de recebimentos e **saídas reais** (nova aba `Saídas de Caixa`)
- Pagamentos sincronizados com o Financeiro (nova aba `Importar Registos`)
- Relatórios sem o limite de 40 reservas
- PDF em pasta própria no Drive, com limpeza automática dos temporários

### Correções aplicadas na revisão desta versão
- **`Pagamentos`:** removidas 351 células com `=#ERR520!` que bloqueavam a expansão do `FILTER`
- **`Dashboard`:** corrigidos os intervalos truncados `$A$3:A3` e `$D$3:D3` (KPIs «Nº de
  reservas» e «Nº de clientes» davam sempre 1, e arrastavam os dois «ticket médio»)
- **`Início` do Financeiro:** instrução atualizada de `A2:AK` para `A2:AF` via `Configurações!B17`

## v2.0 — Modelo relacional

- **`Serviços`** — vários voos/hotéis por reserva (multi-destino), com escalas, duração,
  embarque/desembarque e check-in/check-out
- **`Passageiros`** — manifesto por reserva, com documento, idade à data da viagem e
  aeroporto de partida individual
- Fluxo de **aprovação**: o atendente lança custos, a gestora define a venda
- **`Gestão de Clientes`** (CRM): canceladas com motivo, em orçamento, por aprovar
- **`Calendário`** mensal automático
- **`Configurações`** — listas e regras parametrizadas
- Pesquisa por nome, telemóvel, email ou ID; linha temporal da reserva
- Dashboard executivo com 16 KPIs e destaques

## v1.0 — Base

- Dois ficheiros ligados por `IMPORTRANGE` com separação de confidencialidade
- Reservas, clientes, pacotes, fornecedores
- Pagamentos a prestações (log append-only)
- Contas a receber, custos fixos, comissões, P&L, dashboard
- Formulário com Apps Script e comprovativo em PDF
