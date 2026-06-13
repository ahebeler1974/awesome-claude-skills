# Planilha Universal de Precificação Gastronômica

Planilha de precificação para negócios de **doces/confeitaria, salgados (fritos/assados) e tortas (doces/salgadas)**, com toda a automação rodando em segundo plano.

## Arquivos

- **`Planilha_Universal_Precificacao_Gastronomica.xlsx`** — a planilha final, pronta para uso no Excel ou Google Sheets (Arquivo › Importar). Já vem com um exemplo preenchido (Brigadeiro Gourmet).
- **`build_planilha.py`** — script que gera o `.xlsx` do zero (openpyxl). Permite reproduzir e versionar a planilha.

## Estrutura (5 abas + configuração oculta)

1. **Banco de Dados de Insumos** — cadastro central; calcula o *custo por unidade mínima* (por g/ml/un) com `SEERRO`, já descontando o **Fator de Rendimento** (perda).
2. **Ficha Técnica Universal** — monta a receita com menus suspensos ligados à Aba 1 (`PROCV`); seletor de tipo (Doce/Salgado/Torta).
3. **Custos & Mão de Obra** — converte pró-labore e custos fixos (gás, energia, água, MEI) em **custo por minuto** (mão de obra, forno/fogão e estrutura).
4. **Precificação (Dashboard)** — consolida ingredientes + embalagem + mão de obra + custos fixos; calcula o preço mínimo e simula **canais de venda** (balcão, cartão, iFood) usando o método do divisor `Preço = Custo / (1 − Margem − Taxa)`, que mantém a margem real intacta após a taxa do canal.
5. **Painel Visual** — composição do preço em gráficos (rosca + barras), paleta de cores e regras de formatação condicional.

## Como gerar

```bash
pip install openpyxl
python build_planilha.py
```
