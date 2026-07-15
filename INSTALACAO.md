# Instalação

Tempo estimado: ~20 minutos, uma única vez.

> ⚠️ **Antes de começar:** os dois ficheiros devem ser importados e ficar **na conta Google da
> agência**. No Drive, o dono de um ficheiro é quem o cria — partilhar não transfere a posse.

## 1. Operacional

1. Drive da agência ▸ `Novo ▸ Carregar ficheiro` ▸ `planilhas/Operacional_Agencia_Final.xlsx`.
2. Abrir e usar `Ficheiro ▸ Guardar como Google Sheets` (não deixar em `.xlsx`).
3. `Extensões ▸ Apps Script` ▸ apagar o que lá estiver ▸ colar `src/apps-script/Lets_Go_script.gs` ▸ Guardar.
4. Voltar à folha e atualizar (F5). Deve aparecer o menu **Let's Go**.
5. Correr uma ação do menu e **autorizar** (inclui Google Drive, necessário para o PDF).

## 2. Financeiro

1. Carregar `planilhas/Financeiro_Agencia_Final.xlsx` e converter para Google Sheets.
2. **Não partilhar com ninguém.**
3. Aba `Configurações`:
   - `B16` → ano do relatório (ex.: `2026`)
   - `B17` → URL completo da planilha Operacional
   - `B18` → saldo inicial de caixa
4. Aba `Importar Reservas`: apagar os dados de exemplo em `A3:AF` e colocar em `A3`:
   ```
   =IMPORTRANGE(Configurações!B17;"Reservas!A2:AF")
   ```
5. Aba `Importar Registos`: confirmar a fórmula em `A3` e clicar em **Permitir acesso**.

## 3. Partilha e proteção

| Ficheiro | Partilhar com | Permissão |
|---|---|---|
| Operacional | Atendentes | Editor |
| Financeiro | **Ninguém** | — |

No Operacional, proteger as colunas calculadas (azuis) em `Dados ▸ Proteger intervalos`.

## 4. Pasta dos comprovativos

A gestora cria no Drive da agência a pasta **`Let's Go - Comprovativos`** e partilha-a com os
atendentes como Editores. O script encontra-a pelo nome (inclusive se for partilhada) e grava
lá os PDFs, independentemente de quem carrega no botão.

## 5. Teste de aceitação

Antes de usar a sério, correr o fluxo completo com uma reserva fictícia e confirmar:

- [ ] reserva, cliente e passageiro com IDs únicos
- [ ] custo total, margem, recebido e pendente corretos
- [ ] ações registadas no `Histórico`
- [ ] pagamento visível no Financeiro
- [ ] saída manual refletida no saldo de caixa
- [ ] P&L no ano selecionado
- [ ] PDF gerado na pasta certa

No fim, apagar os dados de exemplo (linhas 2+ de `Clientes`, `Reservas`, `Serviços`,
`Passageiros`, `Registos`).

## ⚠️ Regra permanente

**Nunca exportar e reimportar o sistema em `.xlsx` depois de estar no Google Sheets.** Isso
destrói as fórmulas `FILTER` e os intervalos abertos. Para duplicar, usar sempre
`Ficheiro ▸ Criar uma cópia` dentro do Drive. Ver [docs/licoes-aprendidas.md](docs/licoes-aprendidas.md).
