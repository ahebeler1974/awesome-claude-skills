# Estratégia de Anúncios — Fase 5 (10/06/2026)

## Decisão de plataforma: META ADS

- **Produto:** Ebook Bolos Fit com Marrara Bortoloti (ID 995056) — R$ 49,90, comissão R$ 22,23, CPA máx R$ 15,56
- **Destino dos anúncios:** hotlink Página de Vendas https://go.hotmart.com/Q106261289C
- **Plataforma escolhida:** Meta Ads (Facebook + Instagram)
- **Justificativa (dados):**
  1. **Intenção de busca desfavorável ao Google:** quem busca "bolo fit" / "receita de bolo saudável" quer receita grátis, não comprar ebook. Palavras com intenção de compra ("curso de bolos fit") têm volume baixo e CPC de concorrência de cursos. Com CPA máx de R$ 15,56, Google Search é matematicamente arriscado (CPC PT-BR de nicho culinária/curso R$ 1,50–4,00 → precisaria de conversão ≥ 10–25% na LP, irreal para tráfego frio).
  2. **Produto altamente visual:** bolos = criativo forte em feed/reels. Demonstração visual é o gatilho de compra, não a busca ativa.
  3. **Público claramente segmentável no Meta:** mulheres 25–55, interesses confeitaria/doces fit/renda extra — H1 confirmada como hipótese de partida.
  4. **Conta nova:** setup do Meta é mais simples e rápido que Google (que frequentemente exige aquecimento e verificações em conta zerada).
  5. **LP compatível:** página direta com oferta clara e prova social funciona bem com tráfego de interrupção do Meta (H3).
  6. Nota: Hotmart é certificada no Google Ads (gera link otimizado) — fica registrado como caminho futuro se o Meta falhar no teste.

## Persona principal
- **Quem:** Mulher, 28–55, classe B/C, interessada em vida saudável e doces; muitas já fazem bolo em casa; parte busca renda extra leve com confeitaria.
- **Dor principal:** quer comer doce sem culpa / sem sair da dieta; receitas "fit" da internet dão errado (massa solada, sem sabor).
- **Desejo principal:** bolos gostosos, bonitos e saudáveis que a família come — e, secundariamente, possibilidade de vender.
- **Objeções:** "é só um PDF" / "acho receita grátis no Google" / "receita fit não fica boa" → responder com: 30 receitas testadas por marca com 50 mil alunas, 385 avaliações 4.7★, 5 bônus, garantia de 7 dias.

## Ângulo de venda e compliance
- **Ângulo principal:** "Bolos fit que ficam gostosos de verdade — método testado, não receita de blog."
- **Promessa permitida:** aprender a fazer bolos saudáveis e saborosos com receitas testadas; garantia de 7 dias.
- **Promessas PROIBIDAS (regras nossas + produtor + Meta):**
  - NÃO prometer renda/lucro ("ganhe dinheiro vendendo bolos") — claim de renda viola política do Meta e nossa premissa.
  - NÃO citar diabetes/saúde ("controla diabetes") — claim de saúde.
  - NÃO usar foto/nome/rosto da Marrara como perfil do anúncio (regra do produtor).
  - NÃO usar scarcity falsa ("últimas unidades").
  - NÃO fazer spam (regra Hotmart).

## Estrutura da campanha (proposta para aprovação)
- **1 campanha** — objetivo: Vendas (conversão), orçamento na campanha (CBO) R$ 50/dia.
- **1 conjunto de anúncios:** Brasil, mulheres 25–55, público amplo com sinal de interesse (confeitaria, doces, alimentação saudável) — deixar o Advantage+ expandir; conta nova aprende mais rápido com público amplo.
- **3 criativos (teste A/B/C):**
  - A) Vídeo/carrossel "antes e depois da fatia" — apelo visual do bolo + texto "fit que parece comum".
  - B) Imagem estática com prova social — "385 avaliações ★4.7 / 50 mil alunas".
  - C) Ângulo de dor — "Seu bolo fit fica solado? O problema é a receita, não você."
- **Copies (rascunho):**
  - Título: "Bolos Fit que a família toda come"
  - Corpo: "30 receitas testadas de bolos saudáveis que ficam fofinhos e doces na medida — sem farinha branca e sem açúcar refinado. Método da Marrara Bortoloti com garantia de 7 dias."
  - CTA: "Saiba mais"
- **Pixel/rastreamento:** verificar na Fase 6 a integração de pixel da Hotmart para afiliado (evento Purchase); fallback: otimizar por "Finalização de compra iniciada".

## Metas e limites
- Orçamento: R$ 50/dia (teto absoluto R$ 100/dia) — **nenhum gasto sem aprovação**
- Critério de sucesso: ROAS bruto ≥ 2x E ROI líquido positivo no ciclo de teste (3 vendas/dia = lucro real)
- Critério de pausa: ⚠️ SUBSTITUÍDO pela proposta do Adendo Fase 6 (ver abaixo e fase8 item 3 — pendente aprovação). Original: gasto ≥ R$ 33 sem conversão; CTR < 0,8% após ~1.000 impressões; CPC > R$ 2,00 sustentado.
- Ciclo de teste: 3–5 dias antes de qualquer decisão estrutural

## Riscos
1. Conta Meta nova: CPC instável + risco de bloqueio precoce (mitigar: campanha simples, sem claims, aquecimento com orçamento baixo).
2. Margem apertada: CPA máx R$ 15,56 exige CPC ≤ ~R$ 0,30 com conversão 2% — agressivo; teste dirá.
3. Sem pixel de Purchase confiável no início, a otimização pode ser cega (mitigar: UTMs + relatório Hotmart diário). → RESOLVIDO no Adendo Fase 6.

## Aprovações necessárias (próximas)
1. Aprovar esta estratégia (Fase 5). ✅ Aprovada em 10/06/2026.
2. Criar conta Meta Ads — ✅ criada (act=472207436624172).
3. Aprovar orçamento e publicação (Fase 7). Pendente.

---

# ADENDO FASE 6 (13/06/2026) — Rastreamento, criativos e execução

> Detalhe operacional completo em `fase6-campanha-pronta.md`. Este adendo registra as decisões.

## 1. Rastreamento (resolve o Risco 3)
- **Confirmado:** a Hotmart oferece integração nativa de pixel para afiliados — *Ferramentas → Pixel de Rastreamento* → produto 995056 → Facebook → ID do pixel + token da API de Conversões → evento "compra aprovada" enviado via **Web + API**. O Meta otimiza por **Compra** desde o dia 1.
- **Atribuição por anúncio independente do pixel:** parâmetros no hotlink — `src=metaads&sck=bft1_{{ad.name}}` + UTMs. O `sck` aparece no relatório de vendas da Hotmart → cada venda identifica o anúncio de origem.
- Evento de otimização: **Compra**. Fallback (somente se Purchase não chegar em 48h): "Finalização de compra iniciada".

## 2. Criativos
- Copies finais dos 3 anúncios (A vídeo / B prova social / C dor) prontas no fase6 Parte B3, dentro do compliance.
- Rascunhos B e C gerados no Canva do contratante (links no fase6 A5); vídeo A com roteiro shot-by-shot (gancho nos 2 primeiros segundos).
- Materiais de divulgação do produtor (Hotmart) como fonte preferencial de imagem/vídeo real do produto.

## 3. Identidade de anúncio
- Necessária Página do Facebook (não existia no plano original): nome neutro, sem marca/rosto da Marrara.

## 4. Critérios de pausa — PROPOSTA DE REVISÃO (pendente aprovação)
- Regra original "R$ 33 sem venda" mataria campanhas viáveis durante o aprendizado do algoritmo (falso negativo estimado em 30–50% nos 2 primeiros dias com otimização por Compra).
- Proposta em camadas (detalhe no fase8 item 3): kill de criativo por CTR/CPC; R$ 100 sem "finalização iniciada" → diagnosticar funil; R$ 150 sem venda → trocar produto; **teto total do teste: R$ 250**.

## 5. Expectativa quantificada
- 1ª venda esperada entre R$ 40–120 de gasto com funil saudável; D1 com CPA 2–3× pior é normal.
- Melhor cenário realista deste produto a R$ 50/dia: lucro líquido modesto (R$ 10–25/dia). Função do teste 1: validar a máquina. Lucro escala depois (fase8 item 5: orçamento +20%/48h, criativos semanais, públicos abertos, remarketing com checkout direto, Google Ads plano B, e portfólio com produto de comissão maior).
