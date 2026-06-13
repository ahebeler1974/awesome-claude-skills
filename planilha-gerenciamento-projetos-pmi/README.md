# 📊 Planilha de Gerenciamento de Projetos — PMI / PMBOK

Planilha Excel completa e pronta para uso, alinhada às boas práticas do **PMI (PMBOK)**,
projetada para o dia a dia de um gerente de projetos que precisa de **agilidade**:
o GP preenche o mínimo e a planilha calcula o resto.

**Arquivo:** [`Planilha_Gerenciamento_Projetos_PMI.xlsx`](Planilha_Gerenciamento_Projetos_PMI.xlsx)

## O que ela faz sozinha

| Aba | Automação |
|---|---|
| **📊 Dashboard** | Saúde do projeto (🟢🟡🔴), avanço físico, SPI/CPI, dias para o término, riscos, questões, mudanças, próximos marcos, Curva S e 3 gráficos — tudo calculado das demais abas, nada é digitado |
| **📋 TAP** | Termo de Abertura (Project Charter) com justificativa, objetivo SMART, escopo, premissas, restrições, marcos e aprovações — alimenta o Dashboard |
| **🗂️ EAP** | WBS hierárquica com nível calculado automaticamente a partir do código (1.1.1 → nível 3) |
| **📅 Cronograma** | **Gantt 100% automático**: informe Início, Término e % Concluído — duração, status (Atrasada / Em Andamento / Concluída / Não Iniciada), barras do Gantt, trecho realizado (verde) e destaque da semana atual são calculados sozinhos |
| **💰 EVM** | Gerenciamento de Valor Agregado: informe BAC, % Real e AC — **PV, EV, SV, CV, SPI, CPI, EAC, ETC, VAC e TCPI** são calculados, com semáforo automático e tabela da Curva S |
| **⚠️ Riscos** | Score Probabilidade × Impacto, nível (Baixo → Crítico), exposição financeira e **matriz 5×5 com contagem automática** |
| **👥 Stakeholders** | Quadrante Poder × Interesse calculado (Gerenciar de Perto, Manter Satisfeito, Manter Informado, Monitorar) e alerta quando o engajamento atual difere do desejado |
| **🔄 Mudanças** | Change log com impacto em escopo/prazo/custo e fluxo de aprovação; o custo das mudanças aprovadas soma no Dashboard |
| **🚧 Questões** | Issue log com dias em aberto e alerta 🔴🟡🟢 automáticos por prazo |
| **📚 Lições** | Lições aprendidas por fase e categoria, para registrar durante o projeto |
| **📖 Ajuda** | Guia de uso de cada aba e legenda dos indicadores |

## Recursos

- **Dropdowns** em todos os campos de classificação (status, prioridade, categoria, estratégia de resposta, engajamento…)
- **Formatação condicional** em status, SPI/CPI, score de risco, % concluído e alertas de prazo
- **Painéis congelados e autofiltro** em todas as tabelas
- **Linhas extras pré-formuladas** em todas as abas — basta digitar nas linhas vazias
- Vem preenchida com um **projeto exemplo** (Implantação de ERP) para demonstrar todos os cálculos; basta substituir pelos seus dados

## Regenerar / personalizar

A planilha é gerada por script (Python + openpyxl), o que permite personalizá-la e regenerá-la:

```bash
pip install openpyxl
python3 gerar_planilha.py
```
