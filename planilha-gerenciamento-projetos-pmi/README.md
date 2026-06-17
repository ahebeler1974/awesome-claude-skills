# 📊 Planilha de Gerenciamento de Projetos — PMI + Ágil

Planilha Excel **automatizada, fácil de usar e pronta para o dia a dia** de um gerente de projetos.
Une o rigor do **PMI/PMBOK** com as práticas **ágeis/híbridas** que o mercado usa hoje — e ainda
traz um **Assistente de IA** que transforma o resumo de uma reunião em linhas prontas para a planilha.

**Arquivo:** [`Planilha_Gerenciamento_Projetos_PMI.xlsx`](Planilha_Gerenciamento_Projetos_PMI.xlsx)

---

## ✨ Princípio de uso (o que muda tudo)

> **🟡 Você só preenche os campos AMARELOS. Tudo o que é branco é calculado automaticamente.**

As abas vêm **protegidas (sem senha)** para você não apagar uma fórmula por engano —
só os campos amarelos aceitam digitação. Para liberar tudo: *Revisão → Desproteger Planilha*.

---

## 🤖 O diferencial: Assistente de IA (reunião → planilha)

Acabou a digitação manual depois das reuniões. Na aba **🤖 Assistente IA**:

1. **Cole** o resumo/transcrição da reunião (Teams, Zoom, Meet) na área amarela.
2. **Copie** um dos prompts prontos no Claude/ChatGPT (extrair Ações, Riscos, Decisões ou Questões).
3. **Cole** o resultado na aba indicada — a IA já devolve no formato exato das colunas.

> Usa o Claude/Claude Code? Pule os prompts: cole o texto da reunião na conversa e peça
> *"extraia ações, riscos, decisões e questões no formato da planilha"*.

---

## 🗂️ As 15 abas

| Aba | O que faz sozinha |
|---|---|
| **📊 Dashboard** | Saúde (🟢🟡🔴), avanço físico, SPI/CPI, orçamento, riscos, **sprint atual** e **4 gráficos** — tudo automático |
| **🤖 Assistente IA** | Importa reuniões e devolve linhas prontas para colar (ações/riscos/decisões/questões) |
| **📋 TAP** | Termo de Abertura (Charter): objetivo SMART, escopo, premissas, restrições, marcos, benefícios |
| **🗂️ EAP** | WBS com nível calculado pelo código (1.1.1 → nível 3) |
| **📅 Cronograma** | **Gantt automático**: informe Início, Término e % — duração, status, barras e semana atual são calculados |
| **🏃 Ágil** | Gestão híbrida: backlog com story points, status Kanban, **burndown** do sprint e **velocidade** da equipe |
| **💰 EVM** | **PV, EV, SV, CV, SPI, CPI, EAC, ETC, VAC, TCPI** calculados + Curva S |
| **⚠️ Riscos** | Score Prob.×Impacto, nível, exposição financeira e **matriz 5×5 com contagem automática** |
| **👥 Stakeholders** | Quadrante Poder×Interesse automático + alerta de gap de engajamento |
| **🔄 Mudanças** | Change log com impacto em escopo/prazo/custo |
| **✅ Decisões** | Decision log (artefato moderno, alvo do Assistente de IA) |
| **🚧 Questões** | Issue log com dias em aberto e alerta 🔴🟡🟢 por prazo |
| **📚 Lições** | Lições aprendidas por fase e categoria |
| **📖 Ajuda** | Guia de cada aba e legenda dos indicadores |
| **Listas** | (oculta) alimenta os menus suspensos |

---

## 📈 Gráficos do Dashboard

- **Curva S** — PV × EV × AC acumulados (azul/verde/vermelho)
- **Tarefas por Status** — barras coloridas por situação
- **Riscos por Nível** — pizza (crítico→baixo)
- **Orçamento × Realizado** — BAC vs PV vs EV vs AC

Mais, na aba Ágil: **Burndown do Sprint** e **Velocidade da Equipe**.

---

## ✅ Qualidade / correção

Validação automatizada a cada geração:
- **538 fórmulas** verificadas estaticamente (parênteses, aspas, funções, referências de abas e listas) — **0 erros**
- **Lógica recalculada em Python** contra os dados de exemplo: EVM (BAC = R$ 850.000, CPI, EAC, VAC),
  matriz de riscos (fecha em 10), quadrantes de stakeholders e métricas de sprint — todas as invariantes conferem

> O ambiente de geração não possui Excel/LibreOffice; ao abrir, o Excel recalcula tudo na hora
> (se preciso, tecle **F9**).

---

## 🔧 Regenerar / personalizar

```bash
pip install openpyxl
python3 gerar_planilha.py
```

Vem com um **projeto de exemplo** (Implantação de ERP) demonstrando todos os cálculos —
substitua pelos seus dados nos campos amarelos.
