# Lições aprendidas

Notas honestas do que correu mal e do que ficaria diferente. É a parte do projeto que mais
me ensinou — e a que mais tempo custou a diagnosticar.

---

## 1. O round-trip `.xlsx` ↔ Google Sheets destrói fórmulas modernas

**O sintoma.** Numa versão do sistema, a aba `Pagamentos` tinha **351 células** com o conteúdo
`=#ERR520!` e o Dashboard reportava sempre "1 cliente", independentemente dos dados.

**A causa.** Os ficheiros tinham passado por Excel/LibreOffice. Duas coisas acontecem nessa
viagem:

| No Google Sheets | Depois de gravar em `.xlsx` | Efeito |
|---|---|---|
| `COUNTA('Feed'!$A$3:A)` — intervalo aberto | `COUNTA('Feed'!$A$3:A3)` | Conta **uma** célula → KPI sempre a 1 |
| `FILTER(...)` a expandir por 40 linhas | Linhas por baixo gravadas como `=#ERR520!` | Ao voltar ao Sheets, o `FILTER` não consegue expandir → `#REF!` |

O `.xlsx` não tem o conceito de intervalo aberto nem de *spill* dinâmico do Google, por isso o
conversor "materializa" o que vê — e o que era dinâmico passa a lixo estático.

**A lição.** O `.xlsx` serve para **entregar** o sistema uma vez. A partir do momento em que
está vivo no Sheets, duplica-se dentro do Drive (`Ficheiro ▸ Criar uma cópia`) e nunca mais
se exporta/reimporta.

**Efeito colateral irónico.** O meu próprio pipeline de validação (recalcular em LibreOffice)
tem este defeito embutido: **valida** bem fórmulas clássicas, mas **corrompe** ficheiros com
`FILTER`. Passei a validar esses estaticamente, lendo o XML com `openpyxl` sem gravar.

---

## 2. Locale não é um detalhe

Escrever `=SUMIF(A:A,"x",B:B)` num Google Sheets configurado em português falha: o separador
de argumentos é `;`, não `,`. Ao inserir dados por automação, a solução robusta foi:

- colar **valores** (datas em ISO `aaaa-mm-dd`, números sem separador de milhares);
- **propagar as fórmulas que já existem** na folha com `Ctrl+D`, em vez de escrever fórmulas novas.

Assim é o próprio Sheets que trata da sintaxe e do formato, e o código fica agnóstico ao locale.

---

## 3. "Esconder" nunca é "proteger"

O primeiro instinto foi ocultar as colunas de comissão. Está errado: no Sheets, qualquer
pessoa com edição reexibe uma coluna em dois cliques. A confidencialidade teve de subir um
nível — para a **arquitetura** (dois ficheiros, um deles nunca partilhado).

O mesmo se aplica ao PDF: `DriveApp.createFile()` grava no Drive de **quem executa o script**,
não no da agência. Perceber isto explica porque é que uma pasta partilhada é a solução certa
(`getFoldersByName` também encontra pastas partilhadas).

---

## 4. Um teste de aceitação vale mais do que 100 validações de fórmula

Os livros passavam com "0 erros" no recalculador e mesmo assim tinham bugs de negócio.
O que os apanhou foi correr o fluxo real de ponta a ponta:
criar reserva → adicionar serviço → adicionar passageiro → aprovar → pagar → gerar PDF,
e conferir custo, margem, recebido, pendente, histórico e P&L.

---

## 5. Concorrência aparece mais cedo do que se pensa

Com **dois** atendentes já é possível gerar IDs duplicados: ambos leem o último ID ao mesmo
tempo e ambos escrevem `R-2026-019`. A correção foi `LockService.getDocumentLock()` à volta
das operações que geram IDs, e rollback no `catch` para não deixar uma reserva meio-criada.

---

## 6. O que faria diferente

- **Modelar `Serviços`/`Passageiros` desde o dia 1.** Comecei com colunas fixas de fornecedor
  e tive de reestruturar tudo quando surgiu o multi-destino. O custo da mudança tardia foi
  remapear o Financeiro inteiro — as colunas das reservas mudaram todas de sítio.
- **Perguntar pelo processo, não pelos campos.** O requisito "a gestora aprova o preço" mudou
  o modelo de dados. Se tivesse perguntado *"quem decide o quê?"* na primeira conversa,
  teria aparecido logo.
- **Fixar as regras de negócio numa aba `Configurações` desde o início**, em vez de as ter
  espalhadas por fórmulas.
