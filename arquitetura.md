# Registo de decisões (ADR)

Cada decisão está no formato: contexto → decisão → consequências (boas e más).
O objetivo é deixar registado **porquê**, para que quem pegar nisto daqui a um ano
(inclusive eu) não desfaça uma decisão sem perceber o que ela protegia.

---

## ADR-001 · Google Sheets em vez de Excel

**Contexto.** A agência tem dois atendentes e uma gestora. Precisa de acesso simultâneo,
partilha diferenciada e, no futuro, integração com um bot de WhatsApp.

**Decisão.** Google Sheets.

**Porquê.** A confidencialidade do projeto assenta em **permissões por ficheiro** — algo que
o Sheets faz nativamente e o Excel + OneDrive faz mal. Além disso: `IMPORTRANGE` para ligar
os dois ficheiros, Apps Script para automação sem instalar nada, e edição concorrente real.

**Consequências.**
- ✅ Partilha, automação e concorrência resolvidas de origem.
- ❌ Dependência de ligação à internet e do ecossistema Google.
- ❌ Funções `FILTER`/`IMPORTRANGE`/`COUNTUNIQUE` não existem em `.xlsx` — obriga a cuidado
  na entrega (ver [licoes-aprendidas.md](licoes-aprendidas.md)).

---

## ADR-002 · Dois ficheiros, não um ficheiro com abas ocultas

**Contexto.** Comissões, margens e custos fixos não podem ser vistos pelos atendentes.

**Decisão.** Dois ficheiros: `Operacional` (partilhado) e `Financeiro` (só a gestora),
ligados por `IMPORTRANGE` num único sentido.

**Porquê.** No Google Sheets **as permissões existem ao nível do ficheiro, não da aba**.
Ocultar ou proteger abas é uma barreira cosmética: qualquer pessoa com acesso de edição
pode reexibi-las. A única separação real é não partilhar o ficheiro.

**Consequências.**
- ✅ Confidencialidade estrutural: quem tem o Operacional não tem como chegar ao Financeiro.
- ✅ O `IMPORTRANGE` é só-leitura e não é navegável ao contrário.
- ❌ Dois ficheiros para manter; mudar colunas no Operacional obriga a remapear o Financeiro.
- ⚠️ O Operacional continua a precisar de custo e venda para calcular a reserva. Para
  confidencialidade total dessas colunas seria preciso um terceiro nível de separação.

---

## ADR-003 · Modelo relacional em vez de colunas fixas

**Contexto.** Pedido do cliente: viagens multi-destino (vários voos, vários hotéis) e grupos
em que só uma pessoa paga mas todos precisam de dados para o check-in.

**Decisão.** Duas tabelas novas ligadas por `ID Reserva`:
- **`Serviços`** — uma linha por voo/hotel/transporte (sem limite por reserva).
- **`Passageiros`** — uma linha por pessoa (manifesto).

**Alternativa rejeitada.** Colunas fixas na linha da reserva (`Voo 1`, `Voo 2`, `Hotel 1`…).

**Porquê.** As colunas fixas têm um limite rígido — um roteiro de três cidades já não cabe —
e, com escalas, horários e check-in/out a duplicar, a linha passaria a ter dezenas de colunas
quase sempre vazias. O modelo relacional não tem teto e mantém o preenchimento organizado.

**Consequências.**
- ✅ Multi-cidade sem limite; custo total da reserva = `SUMIF(Serviços)`.
- ✅ Idade e alerta de documento **por passageiro**, e aeroporto de partida individual.
- ❌ A informação de uma reserva fica repartida por várias linhas — resolvido com vistas
  (`Ficha de Reserva`, `Comprovativo`) que reagregam via `FILTER`.

---

## ADR-004 · `PASSAGEIROS` separado de `CLIENTES`

**Contexto.** Numa família de 5, só uma pessoa contrata e paga.

**Decisão.** O **cliente** é quem contrata/paga (catálogo `Clientes`). O **passageiro** é quem
viaja (tabela `Passageiros`, ligada à reserva). O titular aparece nas duas — inserido
automaticamente no manifesto pelo script.

**Porquê.** Tratar os 5 como clientes inflacionaria o catálogo com gente que nunca comprou e
distorceria todas as métricas de negócio (nº de clientes, ticket médio, CRM).

**Consequências.**
- ✅ Métricas fiéis; CRM aponta a quem decide a compra.
- ❌ Se um acompanhante vier a comprar sozinho, os dados são reintroduzidos como cliente novo.

---

## ADR-005 · Aprovação explícita do valor de venda

**Contexto.** Requisito de negócio: o funcionário lança os custos, mas quem define o preço
ao cliente é a gestora.

**Decisão.** A reserva nasce em `Aguarda aprovação` **sem valor de venda**. Só a ação
`Aprovar reserva / definir venda` preenche a venda e passa o estado a `Aprovado`.
O Financeiro só contabiliza receita e custo de reservas aprovadas.

**Consequências.**
- ✅ O preço nunca sai sem passar pela gestão; margem negativa exige confirmação explícita.
- ✅ O P&L não é poluído por orçamentos e cancelamentos.
- ❌ Introduz um passo manual — se a gestora não aprovar, a reserva fica parada
  (mitigado pelo KPI «Por aprovar» e pelo bloco no CRM).

---

## ADR-006 · Pagamentos como log append-only

**Contexto.** Reservas a prestações; era preciso saber quanto falta e quando entrou o quê.

**Decisão.** Cada pagamento é uma linha em `Registos`. A reserva agrega com `SUMIF`/`COUNTIF`.

**Alternativa rejeitada.** Um campo "valor pago" editado à mão.

**Porquê.** Um contador editável perde o histórico e não é auditável. O log dá extrato,
permite conciliação com o banco e alimenta o Financeiro sem lógica adicional.

**Consequências.**
- ✅ Histórico completo; recibos e conciliação possíveis.
- ✅ O Financeiro importa `Registos` e calcula o caixa a partir de recebimentos reais.
- ❌ Corrigir um erro exige uma linha de estorno (não se "edita" o passado).

---

## ADR-007 · `INDEX`/`MATCH` em vez de `XLOOKUP`

**Decisão.** Todas as pesquisas usam `INDEX`/`MATCH`.

**Porquê.** O pipeline de validação recalcula os livros em LibreOffice headless para apanhar
`#REF!`/`#N/A` antes da entrega. O `XLOOKUP` não é suportado nesse motor, o que geraria
falsos positivos e cegaria a validação.

**Consequências.** ✅ Validação automática fiável e compatibilidade ampla. ❌ Fórmulas mais verbosas.
