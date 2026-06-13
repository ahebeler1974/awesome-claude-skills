# FASE 6 — CAMPANHA PRONTA PARA MONTAGEM (atualizado 13/06/2026)

> Este documento contém TUDO o que precisa ser feito para publicar a primeira campanha.
> Parte A = ações que só o contratante pode fazer (15–25 min).
> Parte B = montagem da campanha no Gerenciador de Anúncios (campo a campo, copiar e colar).
> Parte C = protocolo de publicação (aprovação obrigatória antes de qualquer gasto).

---

## PARTE A — Ações pessoais do contratante (pré-requisitos)

### A1. Método de pagamento na conta de anúncios
- Acessar: Gerenciador de Anúncios → Configurações de pagamento (conta **act=472207436624172**)
- Adicionar cartão/método. Sem isso o Meta não publica nada.

### A2. Página do Facebook (identidade do anúncio) — OBRIGATÓRIO
Anúncio no Meta exige uma Página. Se ainda não existe:
- Criar Página simples: nome neutro tipo **"Receitas de Bolo Fit"** ou "Cozinha Fit da [seu nome]"
- Foto de perfil: foto de um bolo (NUNCA foto/nome/rosto da Marrara — regra do produtor)
- Não precisa de seguidores nem posts para anunciar (1–2 posts de receita dão credibilidade, opcional)

### A3. Criar o Pixel no Meta
- Gerenciador de Eventos (business.facebook.com/events_manager2) → Conectar dados → Web → "Criar pixel"
- Nome sugerido: "Pixel Bolos Fit"
- **Copiar o ID do pixel** (número de ~15 dígitos)
- Em Configurações do pixel → API de Conversões → **Gerar token de acesso** → copiar o token

### A4. Integração Hotmart → Meta (o passo que faz o Meta otimizar por COMPRA)
Na Hotmart (app.hotmart.com), como afiliado:
- Menu lateral: **Ferramentas → Ver todas → Pixel de Rastreamento**
- Selecionar o produto **Ebook Bolos Fit (995056)**
- Provedor: **Facebook** → colar o **ID do pixel** (A3)
- Eventos: marcar **Vendas realizadas (compra aprovada)** e, se disponível, **visita ao checkout / finalização de compra iniciada**
- Envio: escolher **Web + API de Conversões** → colar o **token** (A3)
- Salvar
- Referência oficial: help.hotmart.com → artigo "Como configurar o Pixel de Rastreamento" (360015341051) e "Configurações avançadas do pixel" (6565866201741)

Com isso, toda compra aprovada vinda do nosso hotlink dispara o evento **Purchase** no nosso pixel — o Meta passa a otimizar por comprador, não por clicador. Esta é a correção mais importante da missão inteira.

### A5. Criativos
- **Rascunhos prontos no seu Canva** (gerados em 13/06, editáveis — revisar texto/foto antes de usar):
  - **Criativo B (prova social):** https://www.canva.com/d/sdcOf0r0Vm1RN8I
  - **Criativo C (dor):** https://www.canva.com/d/p_YMPUSQT3AblII
  - Candidatos alternativos (escolher outro se preferir):
    - B: https://www.canva.com/d/Mdj_vBCT_r0bRRO | https://www.canva.com/d/N0ymjhEaTFu3sw3 | https://www.canva.com/d/YK91jkP51WLrYjB
    - C: https://www.canva.com/d/uQ64vaPPUUToLeS | https://www.canva.com/d/aRPEam8MDo-A6eP | https://www.canva.com/d/csHigNBD8jjxedU
  - Baixar em PNG (Compartilhar → Baixar) no tamanho 1080×1350.
- **Criativo A (vídeo)** — o mais importante; ver roteiro na Parte B3. Fonte de material: Hotmart → página do produto → **Materiais de divulgação** do produtor (fotos/vídeos oficiais dos bolos), ou gravação própria. PROIBIDO: rosto/nome da Marrara na imagem.
- Conferir em cada peça: sem promessa de renda, sem claim de saúde/diabetes, sem escassez falsa.

---

## PARTE B — Montagem no Gerenciador de Anúncios (campo a campo)

Quem monta: o contratante seguindo este guia (15–20 min), OU uma sessão do Claude in Chrome com permissão em adsmanager.facebook.com (usar o CHECKPOINT-RETOMADA.md).
Regra: montar tudo e **SALVAR COMO RASCUNHO. NÃO PUBLICAR** (Parte C).

### B1. Campanha
| Campo | Valor |
|---|---|
| Objetivo | **Vendas** |
| Nome | `BFT1 - Bolos Fit - Teste 1` |
| Categoria de anúncio especial | Nenhuma |
| Teste A/B | Desligado |
| Orçamento Advantage da campanha (CBO) | **Ligado — R$ 50,00/dia** |

### B2. Conjunto de anúncios
| Campo | Valor |
|---|---|
| Nome | `BFT1-CJ1 - BR F25-55 confeitaria` |
| Local de conversão | **Site** |
| Pixel | Pixel Bolos Fit (criado em A3) |
| Evento de conversão | **Compra** (se o Gerenciador alertar "evento sem atividade", manter Compra mesmo assim — o evento chegará via Hotmart; fallback só se após 48h não registrar nada: trocar para "Finalização de compra iniciada") |
| Localização | Brasil |
| Idade | 25–55 |
| Gênero | Mulheres |
| Público Advantage+ | Ligado, com sugestões de interesse: **Confeitaria, Bolos, Doces, Alimentação saudável, Receitas** |
| Posicionamentos | **Advantage+ (automáticos)** |
| Atribuição | Padrão (clique 7 dias / visualização 1 dia) |

### B3. Anúncios (3) — identidade: Página criada em A2

**URL do site (igual nos 3):**
```
https://go.hotmart.com/Q106261289C
```

**Parâmetros de URL (campo "Parâmetros de URL", igual nos 3):**
```
src=metaads&sck=bft1_{{ad.name}}&utm_source=meta&utm_medium=paid&utm_campaign={{campaign.name}}&utm_content={{ad.name}}
```
> O `sck` aparece na coluna de rastreamento do relatório de vendas da Hotmart → mesmo sem pixel sabemos QUAL anúncio vendeu. (Referência: help.hotmart.com — "Como identificar a origem das minhas vendas" / "O que são parâmetros de campanha".)

---

#### Anúncio A — `a-video` (vídeo, prioridade máxima)
**Roteiro 15–20s (formato 4:5 para feed; versão 9:16 para stories/reels se possível):**
1. **0–2s (gancho):** close de uma fatia sendo cortada, miolo fofinho aparecendo. Texto na tela: "Isso é um bolo FIT."
2. **2–6s:** ingredientes na bancada. Texto: "Sem farinha branca. Sem açúcar refinado."
3. **6–11s:** bolo crescido no forno/desenformando. Texto: "Bolo fit solado? É a receita errada."
4. **11–16s:** Texto: "30 receitas testadas · 4.7★ (385 avaliações)"
5. **16–20s:** bolo finalizado + fatia. Texto: "Toque em Saiba mais"

**Texto principal:**
```
Bolo fit que fica fofinho de verdade — sem farinha branca e sem açúcar refinado. 🍰

São 30 receitas testadas que não solam e não ficam com gosto de "dieta". A família come sem nem perceber que é saudável.

✔️ Método avaliado com 4.7★ (385 avaliações)
✔️ Garantia incondicional de 7 dias

Toque em Saiba mais e veja as receitas.
```
**Título:** `Bolos Fit que a família toda come`
**Descrição:** `30 receitas testadas · Garantia de 7 dias`
**CTA:** Saiba mais

---

#### Anúncio B — `b-prova` (imagem estática — Canva B)
**Texto principal:**
```
Mais de 50 mil alunas já fazem bolos saudáveis que ficam gostosos de verdade.

30 receitas de bolos fit testadas — fofinhos, doces na medida, sem farinha branca e sem açúcar refinado. Método da Marrara Bortoloti, avaliado com 4.7★ por 385 alunas.

Garantia incondicional de 7 dias: não gostou, recebe seu dinheiro de volta.

Toque em Saiba mais.
```
**Título:** `4.7★ — 385 avaliações reais`
**Descrição:** `Sem farinha branca · Sem açúcar refinado`
**CTA:** Saiba mais

---

#### Anúncio C — `c-dor` (imagem estática — Canva C)
**Texto principal:**
```
Seu bolo fit fica solado, seco ou com gosto estranho? O problema não é você — é a receita. 😉

Receita "fit" de blog não é testada. Estas 30 são: medidas exatas, substituições que funcionam e bolo fofinho saindo do forno.

✔️ 4.7★ (385 avaliações)
✔️ Garantia de 7 dias

Toque em Saiba mais e nunca mais erre um bolo.
```
**Título:** `Seu bolo fit ficou solado?`
**Descrição:** `O problema é a receita — use 30 testadas`
**CTA:** Saiba mais

---

### B4. Checklist de compliance (conferir antes de salvar)
- [ ] Nenhuma promessa de renda/lucro
- [ ] Nenhum claim de saúde (diabetes, emagrecimento garantido etc.)
- [ ] Sem foto/rosto da Marrara nos criativos e na Página
- [ ] Sem escassez falsa
- [ ] URL e parâmetros idênticos nos 3 anúncios
- [ ] Tudo salvo como **RASCUNHO**

---

## PARTE C — Protocolo de publicação

1. Campanha montada como rascunho → revisar com este documento ao lado.
2. **APROVAÇÃO NECESSÁRIA** (formato padrão da missão): publicar campanha BFT1 com orçamento R$ 50/dia. Respostas: 1. Aprovado / 2. Não aprovado / 3. Ajustar.
3. Após "Aprovado": publicar, anotar data/hora no campaign-log, conferir no Gerenciador de Eventos (aba "Testar eventos"/visão geral) se os eventos da Hotmart estão chegando.
4. Iniciar **Fase 8**: monitoramento diário conforme `fase8-otimizacao-escala.md`, registros em `daily-performance.markdown`.

## Observações técnicas
- A página de vendas respondeu 403 a acesso automatizado (proteção anti-bot — normal). Última verificação manual OK em 10/06. **Reconferir o hotlink no navegador antes de publicar.**
- Primeiras ~48h = aprendizado do conjunto: CPA instável é esperado. Não mexer na campanha nesse período, salvo violação dos critérios de pausa (ver fase8).
- Nunca editar anúncio ativo (reseta o aprendizado). Para mudar algo: duplicar, editar a cópia, pausar o antigo.
