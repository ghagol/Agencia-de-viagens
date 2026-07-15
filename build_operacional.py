# Geradores dos livros

Scripts Python (`openpyxl`) que constroem os dois `.xlsx` do zero — estrutura, fórmulas,
validações, formatação condicional e dados de exemplo.

| Ficheiro | Papel |
|---|---|
| `brand.py` | Sistema de design: paleta, tipografia e helpers (`paint_header`, `zebra`, `tab`, `fill`) |
| `sample_data.py` | Dados de exemplo (fictícios). **Fonte única** partilhada pelos dois builds, para os ficheiros nunca divergirem |
| `build_operacional.py` | Gera o livro Operacional |
| `build_financeiro.py` | Gera o livro Financeiro |

## Utilização

```bash
pip install openpyxl
python build_operacional.py
python build_financeiro.py
```

## Validação

Durante o desenvolvimento os livros eram recalculados em LibreOffice headless para apanhar
`#REF!`/`#N/A` antes da entrega:

```bash
python recalc.py Operacional_Agencia.xlsx 45
```

> ⚠️ **Atenção:** este passo **não pode** ser usado em livros que contenham `FILTER`. O
> LibreOffice materializa a expansão dessas fórmulas e deixa lixo (`=#ERR520!`) nas células
> por baixo, o que depois bloqueia a expansão no Google Sheets. Livros com `FILTER` devem ser
> validados estaticamente (ler o XML com `openpyxl`, sem gravar).
> Ver [`docs/licoes-aprendidas.md`](../../docs/licoes-aprendidas.md).

## Nota de versões

Estes scripts geram a **base estrutural** do sistema (modelo relacional da v2). As melhorias
da v3 — `LockService`, rollback, histórico com antes/depois, P&L anual, caixa real, abas
`Importar Registos` e `Saídas de Caixa` — foram aplicadas por cima, e estão refletidas nos
livros em [`planilhas/`](../../planilhas) e no script em
[`src/apps-script/`](../apps-script), que são a **fonte de verdade** do que está em produção.
