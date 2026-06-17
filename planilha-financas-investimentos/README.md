# 💰 Planilha de Finanças & Investimentos (PT-BR)

Planilha Excel completa e **validada** para uso real no dia a dia. Reúne
controle financeiro, orçamento, carteira de investimentos, patrimônio,
simulador de juros compostos, independência financeira (FIRE) e metas — tudo
integrado e atualizado automaticamente a partir dos seus lançamentos.

**Arquivo:** `Planilha_Financas_e_Investimentos.xlsx`

## Abas (11)

| Aba | O que faz |
|---|---|
| 🏠 **Início** | Capa com menu de navegação e a regra de ouro de uso. |
| 📊 **Dashboard** | 12 indicadores em cartões + 4 gráficos (receitas × despesas × investimentos por mês, despesas por categoria, evolução do patrimônio e taxa de poupança). Tudo automático. |
| 💰 **Lancamentos** | Controle financeiro com menus suspensos, cores automáticas por tipo, coluna de **Saldo em Caixa** acumulado e 6 meses de dados de exemplo. |
| 📅 **Orcamento** | Orçamento doméstico planejado × realizado por categoria e mês, com alerta de estouro, % do orçamento anual usado (barra de dados), saldo mensal e taxa de poupança. |
| 📈 **Carteira** | Ações, FIIs e ETFs com resultado em R$ e %, **alocação-alvo e rebalanceamento** (quanto comprar/vender), yield sobre custo, **retorno total com dividendos**, resumo por classe e 3 gráficos. |
| 🏦 **Patrimonio** | Patrimônio líquido mês a mês (puxa o total da Carteira), variação, **reserva de emergência** com meta e %, e gráficos de evolução e composição. |
| 🧮 **Simulador** | Juros compostos com 3 cenários editáveis, ajuste por **inflação** (valor de hoje), renda mensal estimada, regra dos 72, e gráficos de linha e área empilhada (aportado × juros). |
| 🔥 **Independencia** | Calculadora FIRE: número da independência, anos e idade até alcançá-la, **Lean / Fat / Coast FIRE**, marcos de 25/50/75/100% e gráfico da trajetória. |
| 🎯 **Metas** | Objetivos com valor alvo, prazo, % concluído e **aporte mensal necessário** para chegar no prazo, com status e barra de progresso. |
| 📖 **Instrucoes** | Guia de uso passo a passo. |
| ⚙️ **Config** | Listas de categorias, tipos, formas de pagamento e classes (personalize os menus suspensos aqui). |

## Como usar

1. Abra no Excel, Google Sheets ou LibreOffice Calc.
2. Comece pela aba **Início**: ela traz um resumo ao vivo e **links clicáveis**
   para todas as abas.
3. Edite **apenas as células amarelas** (🟡 = você preenche) e a aba
   **Lançamentos**. As células cinza (⬜) são cálculo automático.
4. Apague os dados de exemplo (jan–jun/2026) e registre os seus.

## Novidades desta versão (foco em usabilidade profissional)

- **Aba Início interativa**: 8 indicadores ao vivo (patrimônio, receitas,
  despesas, poupança, carteira, FIRE, saldo, reserva) + **menu de navegação
  clicável** para cada aba, e legenda de cores.
- **Botão “🏠 Início”** em todas as abas para voltar ao menu com um clique.
- **Padrão de cores consistente**: 🟡 amarelo = entrada do usuário,
  ⬜ cinza = fórmula automática — em todas as abas.
- **Painel de Insights automáticos** no Dashboard: maior categoria de despesa,
  mês de maior gasto, médias de receitas/despesas/aportes, sobra média,
  dividendos médios e anos até a independência (`INDEX/MATCH`, `AVERAGEIF`,
  `COUNTIF`).
- **Filtro/ordenação** (AutoFiltro) na aba Lançamentos.
- **Abas de exibição protegidas** (Dashboard e Instruções, sem senha) para
  evitar apagar fórmulas por engano — basta *Revisão → Desproteger* para
  liberar.

## Confiabilidade — como esta planilha foi validada

- **2.024 fórmulas**, todas com sintaxe verificada e proteção contra erros
  (`IFERROR`).
- Cada cálculo foi **recalculado por um motor independente** e comparado com
  matemática feita à parte em Python (incl. `numpy-financial` para `FV`/`NPER`):
  **45 verificações, 0 falhas, 0 células de erro** (`#REF!`, `#DIV/0!`, etc.).
- **Referências limitadas** (sem colunas inteiras) para a planilha permanecer
  rápida e estável mesmo com muitos lançamentos.
- 13 gráficos nativos, 17 regras de formatação condicional e 4 menus suspensos.
- **Sem macros (VBA)** — funciona em qualquer Excel moderno, Google Sheets e
  LibreOffice.
