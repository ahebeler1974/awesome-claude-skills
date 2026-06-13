# 💰 Planilha de Finanças & Investimentos (PT-BR)

Planilha Excel completa para uso real no dia a dia, com 5 ferramentas integradas
que se atualizam automaticamente a partir dos seus lançamentos.

**Arquivo:** `Planilha_Financas_e_Investimentos.xlsx`

## Abas

| Aba | O que faz |
|---|---|
| 📊 **Dashboard** | Painel com 8 indicadores (receitas, despesas, taxa de poupança, carteira, progresso FIRE…) e gráficos de receitas × despesas e despesas por categoria. Tudo automático. |
| 💰 **Lancamentos** | Controle financeiro: registre cada entrada e saída com menus suspensos de Tipo, Categoria e Forma de Pagamento. Cores automáticas por tipo. Vem com 6 meses de dados de exemplo. |
| 📅 **Orcamento** | Orçamento doméstico planejado × realizado, mês a mês, por categoria. Destaca em vermelho quando o orçamento estoura e calcula saldo mensal e taxa de poupança. |
| 📈 **Carteira** | Carteira de ações, FIIs e ETFs: preço médio, resultado em R$ e %, participação de cada ativo, dividendos, yield sobre custo e gráfico de composição. |
| 🧮 **Simulador** | Simulador de juros compostos com 3 cenários de rentabilidade editáveis (conservador, moderado, arrojado), projeção de 30 anos e gráfico comparativo. |
| 🔥 **Independencia** | Calculadora FIRE: número da independência financeira, anos até alcançá-la, idade na independência, marcos de 25/50/75/100% e gráfico da trajetória. |
| 📖 **Instrucoes** | Guia de uso passo a passo dentro da própria planilha. |
| ⚙️ **Config** | Listas de categorias, tipos e formas de pagamento (personalize aqui os menus suspensos). |

## Como usar

1. Abra o arquivo no Excel (ou Google Sheets / LibreOffice Calc).
2. Leia a aba **Instrucoes**.
3. Edite **apenas as células amarelas** e a aba **Lancamentos** — todo o resto
   é calculado por fórmulas.
4. Apague os dados de exemplo (jan–jun/2026) e comece a registrar os seus.

## Detalhes técnicos

- Gerada com Python + openpyxl; fórmulas nativas do Excel (`SUMIFS`, `FV`,
  `NPER`, `INDEX`/`MATCH`), validação de dados, formatação condicional,
  barras de dados e 5 gráficos nativos.
- Sem macros (VBA) — funciona em qualquer Excel moderno e no Google Sheets.
