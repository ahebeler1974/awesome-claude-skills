# FASE 8 — OTIMIZAÇÃO, PRIMEIRA VENDA E ESCALA (criado 13/06/2026)

> Objetivo declarado do contratante: **fazer a primeira venda de afiliado com tráfego pago, depois escalar com lucro.**
> Este documento define como operar a campanha do dia 1 até a escala — e o risco máximo aceito.

---

## 1. Por que as tentativas anteriores nunca venderam (diagnóstico honesto)

As 4 causas clássicas — e o que esta missão faz de diferente:

1. **Otimizar por clique, não por compra.** Sem evento Purchase, o Meta entrega o anúncio para quem clica em tudo e não compra nada. → Corrigido: integração Hotmart→Meta (pixel + API de Conversões) faz o Meta caçar COMPRADORAS (fase6, A3–A4).
2. **Matar a campanha cedo demais.** O algoritmo precisa de ~48h+ de aprendizado; quase toda campanha vencedora parece perdedora no dia 1. A regra antiga (pausar com R$ 33 sem venda) tinha chance alta de abortar uma campanha que venderia no dia 2. → Corrigido: regras revisadas abaixo (item 3), com teto de risco total definido.
3. **Criativo fraco para tráfego frio.** No Meta atual, o criativo é ~80% do resultado. → Corrigido: vídeo como prioridade + 2 estáticos com ângulos diferentes, todos com gancho nos 2 primeiros segundos.
4. **Voar às cegas na atribuição.** Sem saber qual anúncio vendeu, otimização vira chute. → Corrigido: `sck=bft1_{{ad.name}}` em todos os anúncios — cada venda na Hotmart mostra o anúncio de origem, independente do pixel.

## 2. Expectativa realista (para não desistir na hora errada)

- Comissão: R$ 22,23 | CPA máximo: R$ 15,56 | Empate: 2,25 vendas/dia (R$ 50/dia)
- **Janela típica da 1ª venda com funil saudável: D1–D3, entre R$ 40 e R$ 120 de gasto.**
- Dia 1 quase sempre tem CPA 2–3× pior que a média da semana (aprendizado). Decisões estruturais só com 3+ dias de dados — exceto violação dos critérios de pausa.
- Este produto, neste orçamento, no melhor cenário realista gera lucro líquido modesto (R$ 10–25/dia). **O papel dele é validar a máquina (1ª venda + CPA viável). O lucro grande vem da escala (item 5).**

## 3. Regras de operação — PROPOSTA REVISADA ⚠️ (requer aprovação)

**APROVAÇÃO NECESSÁRIA** — substitui o critério de pausa "R$ 33 sem venda" do plano da Fase 5:

| Nível | Gatilho | Ação |
|---|---|---|
| Criativo | CTR < 0,8% após ~1.000 impressões | Pausar só o criativo |
| Criativo | CPC > R$ 2,00 sustentado (não nas primeiras horas) | Pausar só o criativo |
| Campanha | R$ 100 gastos (≈2 dias) sem NENHUMA "finalização de compra iniciada" | Pausar campanha e diagnosticar funil (problema = página/oferta, não anúncio) |
| Campanha | R$ 150 gastos sem nenhuma venda | Pausar e decidir troca de produto (backups abaixo) |
| Missão | **Teto total do teste: R$ 250** | Fim do teste 1 — análise completa antes de qualquer real adicional |

- Risco máximo absoluto deste teste: **R$ 250** (5 dias × R$ 50). Nenhum centavo além sem nova aprovação.
- Racional da mudança: com otimização por Compra, a variância dos 2 primeiros dias faz a regra de R$ 33 matar campanhas boas em ~30–50% dos casos. R$ 150 ≈ 6,7× a comissão — se não saiu venda até aí, o problema é estrutural mesmo, não variância.
- Backups de produto já validados (se trocar): Doces Celebrações (R$ 57,90 / com. R$ 20,67), Laço Inquebrável pets (R$ 55,90 / R$ 19,95), Horta hidropônica (R$ 29,90 / R$ 21,15).

## 4. Rotina diária (10 min — registrar em daily-performance.markdown)

1. **Gerenciador de Anúncios:** gasto, impressões, CPM, CTR, CPC, finalizações iniciadas, compras, CPA — total e POR ANÚNCIO.
2. **Hotmart (Vendas):** vendas aprovadas + coluna SCK (qual anúncio vendeu) + valor da comissão (conferir se houve order bump — comissão média pode passar de R$ 22,23).
3. Aplicar tabela do item 3. Em dúvida entre mexer e esperar: **esperar** (mexer reseta aprendizado).
4. Diagnóstico por funil:
   - Impressões altas + CTR baixo → problema de CRIATIVO
   - Cliques ok + zero finalização iniciada → problema de PÁGINA/OFERTA
   - Finalizações iniciadas + zero compra → problema de CHECKOUT/PREÇO (ou atribuição — conferir Hotmart antes de concluir)
5. **Conferir o pixel nos 2 primeiros dias:** Gerenciador de Eventos → eventos Purchase chegando? Se não: revisar integração Hotmart (fase6 A4) e, enquanto isso, confiar no relatório Hotmart + SCK.

## 5. Escada de escala (após a 1ª venda)

**Gatilho de escala: CPA médio ≤ R$ 15,56 por 3 dias consecutivos.**

1. **Vertical (orçamento):** +20% no CBO a cada 48h estáveis: 50 → 60 → 72 → 86 → **R$ 100/dia (teto já aprovado — acima disso, nova aprovação)**. Nunca dobrar de uma vez: salto grande reseta o aprendizado e estoura o CPA.
2. **Criativos (o motor real da escala):** 1–2 variações novas POR SEMANA do anúncio vencedor (trocar o gancho dos 2 primeiros segundos, manter o resto). Pausar perdedores. Nunca editar anúncio ativo — duplicar e editar a cópia.
3. **Horizontal (públicos):** com ≥ ~10 compras registradas no pixel → 2º conjunto Advantage+ TOTALMENTE aberto (sem interesses), mesmo orçamento. Mais adiante (≥ 50–100 compras): Lookalike 1% de compradoras.
4. **Remarketing (margem alta):** público de envolvimento (IG/FB 30 dias) + quem iniciou checkout sem comprar → anúncio direto para o checkout `https://go.hotmart.com/Q106261289C?ap=0309&src=metaads&sck=bft1_rmkt`. R$ 10–15/dia, só quando houver volume.
5. **2º canal:** Google Ads com link certificado da Hotmart ("Gerar link para o Google Ads" na página de hotlinks) — plano B já mapeado na Fase 5.
6. **Portfólio (a verdadeira escala de LUCRO):** comissão de R$ 22 limita o jogo por matemática. Com a máquina validada (pixel quente + criativos vencedores + processo), adicionar 1 produto de comissão R$ 80–150+ para a MESMA audiência (culinária/feminina/renda… respeitando regras de compliance). É aqui que R$ 50/dia de lucro vira R$ 500/dia — não no aumento de orçamento do produto de ticket baixo.

## 6. Fórmulas (conferência rápida)
- CPA = gasto ÷ vendas | Empate: CPA = R$ 22,23 | Meta: CPA ≤ R$ 15,56
- ROI líquido/dia = (vendas × R$ 22,23) − gasto
- ROAS bruto = (vendas × R$ 49,90) ÷ gasto — lembrete: ROAS bruto 2x ainda é ROI líquido NEGATIVO (−R$ 5,54/dia); decisão sempre pelo líquido.
