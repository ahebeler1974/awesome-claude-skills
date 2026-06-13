# -*- coding: utf-8 -*-
"""
Gerador da Planilha de Gerenciamento de Projetos PMI/PMBOK.

Gera um arquivo .xlsx completo, com 12 abas interligadas por formulas:
Dashboard executivo, TAP, EAP, Cronograma com Gantt automatico, EVM
(Gerenciamento de Valor Agregado) com Curva S, Registro de Riscos com
matriz 5x5, Stakeholders, Mudancas, Questoes e Licoes Aprendidas.

Uso: python3 gerar_planilha.py
"""

import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- paleta
AZUL_ESCURO = "1F3864"
AZUL = "2E5FA3"
AZUL_CLARO = "D9E2F3"
AZUL_GANTT = "4472C4"
VERDE = "548235"
VERDE_CLARO = "E2EFDA"
VERDE_GANTT = "70AD47"
AMARELO = "FFC000"
AMARELO_CLARO = "FFF2CC"
LARANJA = "ED7D31"
VERMELHO = "C00000"
VERMELHO_CLARO = "FCE4E4"
CINZA = "808080"
CINZA_CLARO = "F2F2F2"
BRANCO = "FFFFFF"

HOJE = dt.date(2026, 6, 13)  # apenas para gerar dados de exemplo coerentes

FMT_DATA = "DD/MM/YYYY"
FMT_MOEDA = 'R$ #,##0.00'
FMT_MOEDA_K = 'R$ #,##0'
FMT_PCT = "0%"
FMT_NUM2 = "0.00"

fino = Side(style="thin", color="BFBFBF")
BORDA = Border(left=fino, right=fino, top=fino, bottom=fino)


def fill(cor):
    return PatternFill("solid", fgColor=cor)


def estilo_titulo(ws, celula, texto, cor=AZUL_ESCURO, tamanho=16):
    c = ws[celula]
    c.value = texto
    c.font = Font(bold=True, size=tamanho, color=BRANCO)
    c.fill = fill(cor)
    c.alignment = Alignment(horizontal="center", vertical="center")


def cabecalho(ws, linha, colunas, cor=AZUL_ESCURO, altura=28):
    """Escreve a linha de cabecalho de uma tabela."""
    for i, nome in enumerate(colunas, start=1):
        c = ws.cell(row=linha, column=i, value=nome)
        c.font = Font(bold=True, size=10, color=BRANCO)
        c.fill = fill(cor)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA
    ws.row_dimensions[linha].height = altura


def larguras(ws, valores):
    for i, w in enumerate(valores, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def celula(ws, linha, col, valor=None, fmt=None, bold=False, cor_fundo=None,
           cor_fonte="000000", alinh="left", wrap=False, tamanho=10, borda=True):
    c = ws.cell(row=linha, column=col)
    if valor is not None:
        c.value = valor
    if fmt:
        c.number_format = fmt
    c.font = Font(bold=bold, size=tamanho, color=cor_fonte)
    if cor_fundo:
        c.fill = fill(cor_fundo)
    c.alignment = Alignment(horizontal=alinh, vertical="center", wrap_text=wrap)
    if borda:
        c.border = BORDA
    return c


def rotulo_valor(ws, linha, col_rotulo, rotulo, valor, fmt=None, merge_ate=None):
    """Par rotulo/valor usado em formularios (TAP, Dashboard)."""
    celula(ws, linha, col_rotulo, rotulo, bold=True, cor_fundo=AZUL_CLARO)
    c = celula(ws, linha, col_rotulo + 1, valor, fmt=fmt, wrap=True)
    if merge_ate:
        ws.merge_cells(start_row=linha, start_column=col_rotulo + 1,
                       end_row=linha, end_column=merge_ate)
        for cc in range(col_rotulo + 1, merge_ate + 1):
            ws.cell(row=linha, column=cc).border = BORDA
    return c


wb = Workbook()

# ================================================================ LISTAS
# Aba oculta com listas de validacao; expostas como nomes definidos para
# que os dropdowns funcionem em qualquer aba.
ws = wb.active
ws.title = "Listas"
listas = {
    "ListaStatus": ["⚪ Não Iniciada", "🟡 Em Andamento", "✅ Concluída", "🔴 Atrasada", "⏸ Suspensa"],
    "ListaPrioridade": ["🔴 Alta", "🟡 Média", "🟢 Baixa"],
    "ListaSimNao": ["Sim", "Não"],
    "ListaCatRisco": ["Técnico", "Externo", "Organizacional", "Gerenciamento", "Comercial", "Regulatório"],
    "ListaEstrategia": ["Mitigar", "Transferir", "Aceitar", "Eliminar", "Explorar", "Compartilhar"],
    "ListaStatusRisco": ["Aberto", "Em Tratamento", "Materializado", "Encerrado"],
    "ListaStatusMudanca": ["Em Análise", "Aprovada", "Rejeitada", "Adiada", "Implementada"],
    "ListaStatusQuestao": ["Aberta", "Em Tratamento", "Resolvida", "Escalada"],
    "ListaEngajamento": ["Desinformado", "Resistente", "Neutro", "Apoiador", "Líder"],
    "ListaFase": ["Iniciação", "Planejamento", "Execução", "Monitoramento", "Encerramento"],
    "ListaCatLicao": ["Processo", "Pessoas", "Tecnologia", "Fornecedores", "Comunicação", "Riscos"],
    "Lista15": [1, 2, 3, 4, 5],
}
for j, (nome, valores) in enumerate(listas.items(), start=1):
    col = get_column_letter(j)
    ws.cell(row=1, column=j, value=nome).font = Font(bold=True)
    for i, v in enumerate(valores, start=2):
        ws.cell(row=i, column=j, value=v)
    ref = f"Listas!${col}$2:${col}${len(valores) + 1}"
    wb.defined_names.add(DefinedName(nome, attr_text=ref))
ws.sheet_state = "hidden"


def dropdown(ws, nome_lista, faixa):
    dv = DataValidation(type="list", formula1=f"={nome_lista}", allow_blank=True,
                        showErrorMessage=True, errorTitle="Valor inválido",
                        error="Selecione um valor da lista suspensa.")
    ws.add_data_validation(dv)
    dv.add(faixa)
    return dv


# ================================================================ AJUDA
ws = wb.create_sheet("📖 Ajuda")
ws.sheet_properties.tabColor = CINZA
ws.sheet_view.showGridLines = False
larguras(ws, [3, 22, 110])
ws.merge_cells("B2:C2")
estilo_titulo(ws, "B2", "PLANILHA DE GERENCIAMENTO DE PROJETOS — PMI / PMBOK", tamanho=18)
ws.row_dimensions[2].height = 36

guia = [
    ("Dashboard", "Visão executiva automática: saúde do projeto, SPI/CPI, avanço físico, riscos, questões e "
                  "Curva S. NADA é digitado aqui — tudo é calculado das demais abas."),
    ("TAP", "Termo de Abertura do Projeto (Project Charter). Preencha no início: objetivo, escopo, premissas, "
            "restrições, marcos e orçamento. As datas e o nome do projeto alimentam o Dashboard."),
    ("EAP", "Estrutura Analítica do Projeto (WBS). Decomponha as entregas usando códigos hierárquicos "
            "(1, 1.1, 1.1.1...). Use os mesmos códigos no Cronograma e no EVM."),
    ("Cronograma", "Plano de trabalho com Gantt automático: informe Início, Término e % Concluído — a Duração, "
                   "o Status (Atrasada/Em Andamento/Concluída) e as barras do Gantt são calculados sozinhos. "
                   "A semana atual fica destacada no cabeçalho."),
    ("EVM", "Gerenciamento de Valor Agregado: informe o orçamento (BAC), o % real e o custo real (AC) de cada "
            "pacote — PV, EV, SV, CV, SPI, CPI, EAC, ETC, VAC e TCPI são calculados automaticamente."),
    ("Riscos", "Registro de riscos: avalie Probabilidade × Impacto (1 a 5) e a planilha calcula o score, o nível "
               "(Baixo→Crítico) e atualiza a Matriz de Riscos 5×5 ao lado."),
    ("Stakeholders", "Registro de partes interessadas: classifique Poder × Interesse (1 a 5) e a estratégia de "
                     "engajamento (Gerenciar de Perto, Manter Satisfeito...) é sugerida automaticamente."),
    ("Mudanças", "Log de solicitações de mudança com impacto em escopo, prazo e custo, e fluxo de aprovação."),
    ("Questões", "Issue log: pendências e impedimentos do dia a dia, com cálculo automático de dias em aberto "
                 "e alerta visual para itens estourados."),
    ("Lições", "Lições aprendidas, registradas durante o projeto (não só no fim!), por fase e categoria."),
]
celula(ws, 4, 2, "ABA", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
celula(ws, 4, 3, "COMO USAR", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
for i, (aba, txt) in enumerate(guia, start=5):
    celula(ws, i, 2, aba, bold=True, cor_fundo=AZUL_CLARO, alinh="center")
    celula(ws, i, 3, txt, wrap=True)
    ws.row_dimensions[i].height = 34

r = 5 + len(guia) + 1
celula(ws, r, 2, "LEGENDA DE SAÚDE", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
legenda = [
    ("🟢 NO RUMO", "SPI e CPI ≥ 0,95 — projeto dentro do prazo e do orçamento."),
    ("🟡 ATENÇÃO", "SPI ou CPI entre 0,85 e 0,95 — desvios moderados, agir preventivamente."),
    ("🔴 CRÍTICO", "SPI ou CPI < 0,85 — desvio relevante, exige plano de recuperação."),
]
for i, (s, txt) in enumerate(legenda, start=r + 1):
    celula(ws, i, 2, s, bold=True, alinh="center")
    celula(ws, i, 3, txt, wrap=True)

# ================================================================== TAP
ws = wb.create_sheet("📋 TAP")
ws.sheet_properties.tabColor = AZUL
ws.sheet_view.showGridLines = False
larguras(ws, [3, 26, 30, 30, 30])
ws.merge_cells("B2:E2")
estilo_titulo(ws, "B2", "TERMO DE ABERTURA DO PROJETO (PROJECT CHARTER)")
ws.row_dimensions[2].height = 30

rotulo_valor(ws, 4, 2, "Nome do Projeto", "Implantação do Sistema ERP Corporativo", merge_ate=5)
rotulo_valor(ws, 5, 2, "Código do Projeto", "PRJ-2026-014", merge_ate=3)
rotulo_valor(ws, 6, 2, "Gerente do Projeto", "Ana Beatriz Souza", merge_ate=3)
rotulo_valor(ws, 7, 2, "Patrocinador", "Carlos Mendes (Diretor de Operações)", merge_ate=3)
rotulo_valor(ws, 8, 2, "Data de Início", dt.date(2026, 3, 2), fmt=FMT_DATA)
rotulo_valor(ws, 9, 2, "Data de Término", dt.date(2026, 10, 30), fmt=FMT_DATA)
rotulo_valor(ws, 10, 2, "Orçamento Aprovado", 850000, fmt=FMT_MOEDA)
celula(ws, 8, 4, "Cliente / Área Demandante", bold=True, cor_fundo=AZUL_CLARO)
celula(ws, 8, 5, "Diretoria de Operações")
celula(ws, 9, 4, "Prioridade Estratégica", bold=True, cor_fundo=AZUL_CLARO)
celula(ws, 9, 5, "🔴 Alta")
celula(ws, 10, 4, "Reserva de Contingência", bold=True, cor_fundo=AZUL_CLARO)
celula(ws, 10, 5, 85000, fmt=FMT_MOEDA)

secoes = [
    ("JUSTIFICATIVA DO PROJETO",
     "Os sistemas legados não se integram, geram retrabalho de ~1.200 h/mês e impedem o fechamento contábil "
     "em menos de 12 dias. A implantação do ERP unifica os processos de finanças, suprimentos e vendas, com "
     "payback estimado em 18 meses."),
    ("OBJETIVO (SMART)",
     "Implantar o ERP nos módulos Financeiro, Suprimentos e Vendas para 320 usuários, até 30/10/2026, com "
     "orçamento de R$ 850.000, reduzindo o tempo de fechamento contábil de 12 para 4 dias."),
    ("ESCOPO INCLUÍDO",
     "• Módulos Financeiro, Suprimentos e Vendas\n• Migração de dados dos últimos 5 anos\n"
     "• Treinamento de 320 usuários\n• 60 dias de operação assistida"),
    ("ESCOPO EXCLUÍDO (FORA DO ESCOPO)",
     "• Módulos de RH e Manufatura (fase 2)\n• Integração com e-commerce\n• Customizações além de 400 h"),
    ("PREMISSAS",
     "• Key users disponíveis 50% do tempo durante a parametrização\n• Infraestrutura cloud contratada até "
     "abril/2026\n• Dados legados entregues pela TI com qualidade mínima acordada"),
    ("RESTRIÇÕES",
     "• Go-live impreterivelmente antes do fechamento anual (nov/2026)\n• Orçamento limitado a R$ 850.000 + "
     "10% de contingência\n• Equipe interna limitada a 8 pessoas dedicadas"),
    ("RISCOS DE ALTO NÍVEL",
     "• Qualidade dos dados legados abaixo do esperado\n• Indisponibilidade de key users\n"
     "• Atraso na contratação da infraestrutura cloud"),
    ("CRITÉRIOS DE SUCESSO",
     "• Go-live até 30/10/2026 com desvio de custo ≤ 5%\n• Fechamento contábil em ≤ 4 dias no 2º mês\n"
     "• Satisfação dos usuários ≥ 80% na pesquisa pós-implantação"),
]
r = 12
for titulo, txt in secoes:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    celula(ws, r, 2, titulo, bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
    ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=5)
    celula(ws, r + 1, 2, txt, wrap=True)
    ws.row_dimensions[r + 1].height = 16 * (txt.count("\n") + 2)
    for cc in range(2, 6):
        ws.cell(row=r + 1, column=cc).border = BORDA
    r += 3

ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
celula(ws, r, 2, "MARCOS PRINCIPAIS", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
cabecalho(ws, r + 1, ["", "Marco", "Data Prevista", "Data Real", "Status"], cor=AZUL)
marcos = [
    ("Kick-off e TAP aprovado", dt.date(2026, 3, 13), dt.date(2026, 3, 13), "✅ Concluída"),
    ("Plano do projeto aprovado", dt.date(2026, 4, 10), dt.date(2026, 4, 14), "✅ Concluída"),
    ("Parametrização concluída", dt.date(2026, 7, 24), None, "🟡 Em Andamento"),
    ("Migração de dados homologada", dt.date(2026, 8, 28), None, "⚪ Não Iniciada"),
    ("Go-live", dt.date(2026, 10, 9), None, "⚪ Não Iniciada"),
    ("Encerramento do projeto", dt.date(2026, 10, 30), None, "⚪ Não Iniciada"),
]
for i, (m, dp, dr_, st) in enumerate(marcos, start=r + 2):
    celula(ws, i, 2, m)
    celula(ws, i, 3, dp, fmt=FMT_DATA, alinh="center")
    celula(ws, i, 4, dr_, fmt=FMT_DATA, alinh="center")
    celula(ws, i, 5, st, alinh="center")
r = r + 2 + len(marcos) + 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
celula(ws, r, 2, "APROVAÇÕES", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
cabecalho(ws, r + 1, ["", "Papel", "Nome", "Data", "Assinatura"], cor=AZUL)
for i, (papel, nome) in enumerate([("Patrocinador", "Carlos Mendes"),
                                   ("Gerente do Projeto", "Ana Beatriz Souza"),
                                   ("Cliente / Demandante", "Diretoria de Operações")], start=r + 2):
    celula(ws, i, 2, papel, bold=True, cor_fundo=AZUL_CLARO)
    celula(ws, i, 3, nome)
    celula(ws, i, 4, None, fmt=FMT_DATA)
    celula(ws, i, 5, None)

# ================================================================== EAP
ws = wb.create_sheet("🗂️ EAP")
ws.sheet_properties.tabColor = AZUL
larguras(ws, [10, 8, 55, 22, 22, 40])
ws.merge_cells("A1:F1")
estilo_titulo(ws, "A1", "EAP — ESTRUTURA ANALÍTICA DO PROJETO (WBS)")
ws.row_dimensions[1].height = 30
cabecalho(ws, 3, ["Código EAP", "Nível", "Entrega / Pacote de Trabalho", "Responsável",
                  "Critério de Aceitação", "Descrição"])
eap = [
    ("1", "Implantação do Sistema ERP Corporativo", "Ana B. Souza", "", "Projeto completo"),
    ("1.1", "Gerenciamento do Projeto", "Ana B. Souza", "Planos aprovados", "Iniciação, planos, monitoramento e encerramento"),
    ("1.1.1", "Termo de Abertura (TAP)", "Ana B. Souza", "Assinado pelo patrocinador", ""),
    ("1.1.2", "Plano de Gerenciamento do Projeto", "Ana B. Souza", "Aprovado pelo comitê", ""),
    ("1.2", "Infraestrutura", "Roberto Lima", "Ambiente disponível e testado", "Cloud, ambientes e acessos"),
    ("1.2.1", "Contratação do ambiente cloud", "Roberto Lima", "Contrato assinado", ""),
    ("1.2.2", "Configuração dos ambientes (DEV/QA/PRD)", "Roberto Lima", "Checklist de infra 100%", ""),
    ("1.3", "Parametrização do ERP", "Fernanda Costa", "Processos validados pelos key users", "Configuração dos módulos"),
    ("1.3.1", "Módulo Financeiro", "Fernanda Costa", "Cenários de teste aprovados", ""),
    ("1.3.2", "Módulo Suprimentos", "Diego Martins", "Cenários de teste aprovados", ""),
    ("1.3.3", "Módulo Vendas", "Juliana Reis", "Cenários de teste aprovados", ""),
    ("1.4", "Migração de Dados", "Roberto Lima", "Carga homologada com ≤ 0,5% de erro", "Extração, limpeza, carga e validação"),
    ("1.4.1", "Mapeamento e limpeza dos dados", "Roberto Lima", "De-para aprovado", ""),
    ("1.4.2", "Cargas de teste e homologação", "Roberto Lima", "3 cargas com sucesso", ""),
    ("1.5", "Testes Integrados", "Juliana Reis", "0 defeitos críticos abertos", "TIU e testes de aceitação"),
    ("1.6", "Treinamento", "Fernanda Costa", "320 usuários treinados, nota ≥ 8", "Material e turmas de capacitação"),
    ("1.7", "Go-live e Operação Assistida", "Ana B. Souza", "Operação estável por 60 dias", "Virada, suporte e estabilização"),
    ("1.8", "Encerramento", "Ana B. Souza", "Termo de encerramento assinado", "Lições aprendidas e transição"),
]
for i, (cod, nome, resp, crit, desc) in enumerate(eap, start=4):
    nivel = cod.count(".") + 1
    celula(ws, i, 1, cod, bold=nivel <= 2, alinh="center",
           cor_fundo=AZUL_CLARO if nivel == 1 else (CINZA_CLARO if nivel == 2 else None))
    celula(ws, i, 2, f"=LEN(A{i})-LEN(SUBSTITUTE(A{i},\".\",\"\"))+1", alinh="center")
    c = celula(ws, i, 3, ("    " * (nivel - 1)) + nome, bold=nivel <= 2,
               cor_fundo=AZUL_CLARO if nivel == 1 else (CINZA_CLARO if nivel == 2 else None))
    celula(ws, i, 4, resp, alinh="center")
    celula(ws, i, 5, crit, wrap=True)
    celula(ws, i, 6, desc, wrap=True)
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:F{3 + len(eap)}"

# ============================================================ CRONOGRAMA
ws = wb.create_sheet("📅 Cronograma")
ws.sheet_properties.tabColor = VERDE
COL_GANTT0 = 13  # coluna M
N_SEMANAS = 36
SEMANA0 = dt.date(2026, 3, 2)  # segunda-feira da 1a semana

ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
estilo_titulo(ws, "A1", "CRONOGRAMA DO PROJETO — GANTT AUTOMÁTICO")
ws.row_dimensions[1].height = 30
celula(ws, 2, 1, "Status, duração e barras do Gantt são 100% automáticos: preencha apenas Início, Término e % Concluído.",
       borda=False, tamanho=9, cor_fonte=CINZA)
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)

cols = ["ID", "EAP", "Tarefa / Entrega", "Responsável", "Prioridade", "Início", "Término",
        "Duração (dias)", "% Concl.", "Status (auto)", "Predecessoras", "Marco?"]
cabecalho(ws, 4, cols)
larguras(ws, [5, 7, 42, 16, 11, 11, 11, 9, 8, 16, 9, 7])
for k in range(N_SEMANAS):
    col = COL_GANTT0 + k
    ws.column_dimensions[get_column_letter(col)].width = 3.2
    d = SEMANA0 + dt.timedelta(weeks=k)
    c = ws.cell(row=4, column=col, value=d)
    c.number_format = "DD/MM"
    c.font = Font(bold=True, size=8, color=BRANCO)
    c.fill = fill(AZUL_ESCURO)
    c.alignment = Alignment(horizontal="center", vertical="bottom", textRotation=90)
    c.border = BORDA
ws.row_dimensions[4].height = 42

# (tarefa, resp, prioridade, inicio, fim, %concluido, pred, marco) — None em % = formula só
tarefas = [
    ("1.1.1", "Elaborar e aprovar o TAP", "Ana B. Souza", "🔴 Alta", dt.date(2026, 3, 2), dt.date(2026, 3, 13), 1.0, "", "Não"),
    ("1.1.1", "🚩 MARCO: Kick-off do projeto", "Ana B. Souza", "🔴 Alta", dt.date(2026, 3, 13), dt.date(2026, 3, 13), 1.0, "1", "Sim"),
    ("1.1.2", "Plano de gerenciamento do projeto", "Ana B. Souza", "🔴 Alta", dt.date(2026, 3, 16), dt.date(2026, 4, 10), 1.0, "2", "Não"),
    ("1.3", "Levantamento de processos (AS-IS/TO-BE)", "Fernanda Costa", "🔴 Alta", dt.date(2026, 3, 23), dt.date(2026, 4, 24), 1.0, "2", "Não"),
    ("1.2.1", "Contratação do ambiente cloud", "Roberto Lima", "🔴 Alta", dt.date(2026, 4, 1), dt.date(2026, 4, 30), 1.0, "3", "Não"),
    ("1.2.2", "Configuração dos ambientes DEV/QA/PRD", "Roberto Lima", "🟡 Média", dt.date(2026, 5, 4), dt.date(2026, 5, 29), 1.0, "5", "Não"),
    ("1.3.1", "Parametrização — Módulo Financeiro", "Fernanda Costa", "🔴 Alta", dt.date(2026, 5, 4), dt.date(2026, 7, 10), 0.65, "4;6", "Não"),
    ("1.3.2", "Parametrização — Módulo Suprimentos", "Diego Martins", "🔴 Alta", dt.date(2026, 5, 11), dt.date(2026, 7, 17), 0.55, "4;6", "Não"),
    ("1.3.3", "Parametrização — Módulo Vendas", "Juliana Reis", "🟡 Média", dt.date(2026, 5, 18), dt.date(2026, 7, 24), 0.45, "4;6", "Não"),
    ("1.4.1", "Mapeamento e limpeza dos dados legados", "Roberto Lima", "🔴 Alta", dt.date(2026, 5, 11), dt.date(2026, 6, 5), 0.80, "6", "Não"),
    ("1.6", "Desenvolvimento do material de treinamento", "Fernanda Costa", "🟢 Baixa", dt.date(2026, 6, 1), dt.date(2026, 7, 31), 0.20, "7", "Não"),
    ("1.3", "🚩 MARCO: Parametrização concluída", "Fernanda Costa", "🔴 Alta", dt.date(2026, 7, 24), dt.date(2026, 7, 24), 0.0, "7;8;9", "Sim"),
    ("1.4.2", "Cargas de teste e homologação dos dados", "Roberto Lima", "🔴 Alta", dt.date(2026, 6, 15), dt.date(2026, 8, 21), 0.10, "10", "Não"),
    ("1.4.2", "🚩 MARCO: Migração homologada", "Roberto Lima", "🔴 Alta", dt.date(2026, 8, 28), dt.date(2026, 8, 28), 0.0, "13", "Sim"),
    ("1.5", "Testes integrados (TIU)", "Juliana Reis", "🔴 Alta", dt.date(2026, 8, 3), dt.date(2026, 9, 11), 0.0, "12;14", "Não"),
    ("1.5", "Testes de aceitação do usuário (UAT)", "Juliana Reis", "🔴 Alta", dt.date(2026, 9, 14), dt.date(2026, 9, 25), 0.0, "15", "Não"),
    ("1.6", "Treinamento dos usuários finais (320)", "Fernanda Costa", "🟡 Média", dt.date(2026, 9, 14), dt.date(2026, 10, 2), 0.0, "11;15", "Não"),
    ("1.7", "Preparação da virada (cutover plan)", "Ana B. Souza", "🔴 Alta", dt.date(2026, 9, 21), dt.date(2026, 10, 2), 0.0, "16", "Não"),
    ("1.7", "🚩 MARCO: Go-live", "Ana B. Souza", "🔴 Alta", dt.date(2026, 10, 9), dt.date(2026, 10, 9), 0.0, "17;18", "Sim"),
    ("1.7", "Operação assistida (hypercare)", "Equipe ERP", "🔴 Alta", dt.date(2026, 10, 12), dt.date(2026, 10, 30), 0.0, "19", "Não"),
    ("1.8", "Lições aprendidas e termo de encerramento", "Ana B. Souza", "🟡 Média", dt.date(2026, 10, 26), dt.date(2026, 10, 30), 0.0, "20", "Não"),
]
L0 = 5
for i, (eapc, tarefa, resp, prio, ini, fim, pct, pred, marco) in enumerate(tarefas):
    r = L0 + i
    celula(ws, r, 1, i + 1, alinh="center")
    celula(ws, r, 2, eapc, alinh="center")
    celula(ws, r, 3, tarefa, bold=marco == "Sim", wrap=True)
    celula(ws, r, 4, resp, alinh="center")
    celula(ws, r, 5, prio, alinh="center")
    celula(ws, r, 6, ini, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 7, fim, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 8, f"=IF(OR(F{r}=\"\",G{r}=\"\"),\"\",G{r}-F{r}+1)", alinh="center")
    celula(ws, r, 9, pct, fmt=FMT_PCT, alinh="center")
    celula(ws, r, 10,
           f"=IF(F{r}=\"\",\"\",IF(I{r}>=1,\"✅ Concluída\","
           f"IF(TODAY()>G{r},\"🔴 Atrasada\",IF(TODAY()>=F{r},\"🟡 Em Andamento\",\"⚪ Não Iniciada\"))))",
           alinh="center")
    celula(ws, r, 11, pred, alinh="center")
    celula(ws, r, 12, marco, alinh="center")
    for k in range(N_SEMANAS):
        ws.cell(row=r, column=COL_GANTT0 + k).border = BORDA
LFIM = L0 + len(tarefas) - 1
# 30 linhas extras prontas para uso, com as mesmas formulas
for r in range(LFIM + 1, LFIM + 31):
    celula(ws, r, 8, f"=IF(OR(F{r}=\"\",G{r}=\"\"),\"\",G{r}-F{r}+1)", alinh="center")
    celula(ws, r, 9, None, fmt=FMT_PCT, alinh="center")
    celula(ws, r, 10,
           f"=IF(F{r}=\"\",\"\",IF(I{r}>=1,\"✅ Concluída\","
           f"IF(TODAY()>G{r},\"🔴 Atrasada\",IF(TODAY()>=F{r},\"🟡 Em Andamento\",\"⚪ Não Iniciada\"))))",
           alinh="center")
    for cc in list(range(1, 8)) + [11, 12] + [COL_GANTT0 + k for k in range(N_SEMANAS)]:
        ws.cell(row=r, column=cc).border = BORDA
LMAX = LFIM + 30

ws.freeze_panes = "F5"
ws.auto_filter.ref = f"A4:L{LFIM}"
dropdown(ws, "ListaPrioridade", f"E{L0}:E{LMAX}")
dropdown(ws, "ListaSimNao", f"L{L0}:L{LMAX}")

g1 = get_column_letter(COL_GANTT0)
g2 = get_column_letter(COL_GANTT0 + N_SEMANAS - 1)
faixa_gantt = f"{g1}{L0}:{g2}{LMAX}"
# 1) trecho ja realizado (verde) 2) barra planejada (azul) — a ordem importa
ws.conditional_formatting.add(faixa_gantt, FormulaRule(
    formula=[f"AND($F{L0}<>\"\",{g1}$4>=$F{L0},{g1}$4<=$G{L0}+6,{g1}$4<=$F{L0}+($G{L0}-$F{L0})*$I{L0})"],
    fill=fill(VERDE_GANTT), stopIfTrue=True))
ws.conditional_formatting.add(faixa_gantt, FormulaRule(
    formula=[f"AND($F{L0}<>\"\",{g1}$4>=$F{L0},{g1}$4<=$G{L0}+6)"],
    fill=fill(AZUL_GANTT), stopIfTrue=True))
# destaque da semana atual no cabecalho
ws.conditional_formatting.add(f"{g1}4:{g2}4", FormulaRule(
    formula=[f"AND(TODAY()>={g1}$4,TODAY()<{g1}$4+7)"], fill=fill(VERMELHO), stopIfTrue=True))
# status colorido
faixa_status = f"J{L0}:J{LMAX}"
ws.conditional_formatting.add(faixa_status, FormulaRule(formula=[f'NOT(ISERROR(SEARCH("Atrasada",J{L0})))'], fill=fill(VERMELHO_CLARO)))
ws.conditional_formatting.add(faixa_status, FormulaRule(formula=[f'NOT(ISERROR(SEARCH("Concluída",J{L0})))'], fill=fill(VERDE_CLARO)))
ws.conditional_formatting.add(faixa_status, FormulaRule(formula=[f'NOT(ISERROR(SEARCH("Andamento",J{L0})))'], fill=fill(AMARELO_CLARO)))
# barra de progresso na coluna %
ws.conditional_formatting.add(f"I{L0}:I{LMAX}", ColorScaleRule(
    start_type="num", start_value=0, start_color="F8696B",
    mid_type="num", mid_value=0.5, mid_color="FFEB84",
    end_type="num", end_value=1, end_color="63BE7B"))

CRONO_STATUS = f"'📅 Cronograma'!$J${L0}:$J${LMAX}"
CRONO_PCT = f"'📅 Cronograma'!$I${L0}:$I${LFIM}"

# ================================================================== EVM
ws = wb.create_sheet("💰 EVM")
ws.sheet_properties.tabColor = AMARELO
ws.merge_cells("A1:R1")
estilo_titulo(ws, "A1", "EVM — GERENCIAMENTO DE VALOR AGREGADO (EARNED VALUE MANAGEMENT)")
ws.row_dimensions[1].height = 30
celula(ws, 2, 1, "Preencha apenas: BAC, % Real e AC. Tudo o mais (PV, EV, SV, CV, SPI, CPI, EAC, ETC, VAC, TCPI) é calculado.",
       borda=False, tamanho=9, cor_fonte=CINZA)
ws.merge_cells("A2:R2")

cols = ["EAP", "Pacote de Trabalho", "Início", "Término", "BAC\n(Orçamento)", "% Planej.\n(auto)",
        "% Real", "PV (Valor\nPlanejado)", "EV (Valor\nAgregado)", "AC (Custo\nReal)",
        "SV", "CV", "SPI", "CPI", "EAC", "ETC", "VAC", "TCPI"]
cabecalho(ws, 4, cols, altura=42)
larguras(ws, [7, 34, 11, 11, 13, 9, 8, 13, 13, 13, 12, 12, 7, 7, 13, 13, 12, 7])

pacotes = [
    ("1.1", "Gerenciamento do Projeto", dt.date(2026, 3, 2), dt.date(2026, 10, 30), 95000, 0.46, 41000),
    ("1.2", "Infraestrutura", dt.date(2026, 4, 1), dt.date(2026, 5, 29), 120000, 1.00, 131500),
    ("1.3", "Parametrização do ERP", dt.date(2026, 5, 4), dt.date(2026, 7, 24), 260000, 0.55, 158000),
    ("1.4", "Migração de Dados", dt.date(2026, 5, 11), dt.date(2026, 8, 28), 110000, 0.35, 36500),
    ("1.5", "Testes Integrados", dt.date(2026, 8, 3), dt.date(2026, 9, 25), 90000, 0.00, 0),
    ("1.6", "Treinamento", dt.date(2026, 6, 1), dt.date(2026, 10, 2), 85000, 0.10, 7200),
    ("1.7", "Go-live e Operação Assistida", dt.date(2026, 9, 21), dt.date(2026, 10, 30), 75000, 0.00, 0),
    ("1.8", "Encerramento", dt.date(2026, 10, 26), dt.date(2026, 10, 30), 15000, 0.00, 0),
]
E0 = 5
for i, (cod, nome, ini, fim, bac, pct_real, ac) in enumerate(pacotes):
    r = E0 + i
    celula(ws, r, 1, cod, alinh="center")
    celula(ws, r, 2, nome)
    celula(ws, r, 3, ini, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 4, fim, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 5, bac, fmt=FMT_MOEDA_K)
    celula(ws, r, 6, f"=IF(C{r}=\"\",\"\",MAX(0,MIN(1,(TODAY()-C{r}+1)/(D{r}-C{r}+1))))", fmt=FMT_PCT, alinh="center")
    celula(ws, r, 7, pct_real, fmt=FMT_PCT, alinh="center", cor_fundo=AMARELO_CLARO)
    celula(ws, r, 8, f"=IF(E{r}=\"\",\"\",E{r}*F{r})", fmt=FMT_MOEDA_K)
    celula(ws, r, 9, f"=IF(E{r}=\"\",\"\",E{r}*G{r})", fmt=FMT_MOEDA_K)
    celula(ws, r, 10, ac, fmt=FMT_MOEDA_K, cor_fundo=AMARELO_CLARO)
    celula(ws, r, 11, f"=IF(E{r}=\"\",\"\",I{r}-H{r})", fmt=FMT_MOEDA_K)
    celula(ws, r, 12, f"=IF(E{r}=\"\",\"\",I{r}-J{r})", fmt=FMT_MOEDA_K)
    celula(ws, r, 13, f"=IFERROR(I{r}/H{r},\"—\")", fmt=FMT_NUM2, alinh="center")
    celula(ws, r, 14, f"=IFERROR(I{r}/J{r},\"—\")", fmt=FMT_NUM2, alinh="center")
    celula(ws, r, 15, f"=IF(E{r}=\"\",\"\",IFERROR(E{r}/N{r},E{r}))", fmt=FMT_MOEDA_K)
    celula(ws, r, 16, f"=IF(E{r}=\"\",\"\",O{r}-J{r})", fmt=FMT_MOEDA_K)
    celula(ws, r, 17, f"=IF(E{r}=\"\",\"\",E{r}-O{r})", fmt=FMT_MOEDA_K)
    celula(ws, r, 18, f"=IFERROR((E{r}-I{r})/(E{r}-J{r}),\"—\")", fmt=FMT_NUM2, alinh="center")
EF = E0 + len(pacotes) - 1
TR = EF + 1  # linha de totais
celula(ws, TR, 1, "", cor_fundo=AZUL_ESCURO)
celula(ws, TR, 2, "TOTAL DO PROJETO", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO)
for cc in (3, 4):
    celula(ws, TR, cc, "", cor_fundo=AZUL_ESCURO)
for cc, letra in [(5, "E"), (8, "H"), (9, "I"), (10, "J"), (11, "K"), (12, "L"), (15, "O"), (16, "P"), (17, "Q")]:
    celula(ws, TR, cc, f"=SUM({letra}{E0}:{letra}{EF})", fmt=FMT_MOEDA_K, bold=True,
           cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO)
celula(ws, TR, 6, f"=H{TR}/E{TR}", fmt=FMT_PCT, bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
celula(ws, TR, 7, f"=I{TR}/E{TR}", fmt=FMT_PCT, bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
celula(ws, TR, 13, f"=IFERROR(I{TR}/H{TR},\"—\")", fmt=FMT_NUM2, bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
celula(ws, TR, 14, f"=IFERROR(I{TR}/J{TR},\"—\")", fmt=FMT_NUM2, bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
celula(ws, TR, 18, f"=IFERROR((E{TR}-I{TR})/(E{TR}-J{TR}),\"—\")", fmt=FMT_NUM2, bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
ws.freeze_panes = "C5"

# SPI/CPI semaforizados
for faixa in (f"M{E0}:N{EF}",):
    ws.conditional_formatting.add(faixa, CellIsRule(operator="lessThan", formula=["0.85"], fill=fill(VERMELHO_CLARO)))
    ws.conditional_formatting.add(faixa, CellIsRule(operator="between", formula=["0.85", "0.9499"], fill=fill(AMARELO_CLARO)))
    ws.conditional_formatting.add(faixa, CellIsRule(operator="greaterThanOrEqual", formula=["0.95"], fill=fill(VERDE_CLARO)))
for faixa in (f"K{E0}:L{EF}", f"Q{E0}:Q{EF}"):
    ws.conditional_formatting.add(faixa, CellIsRule(operator="lessThan", formula=["0"], fill=fill(VERMELHO_CLARO)))

# glossario rapido
gl = TR + 2
celula(ws, gl, 2, "PV = quanto deveria ter sido feito | EV = quanto foi feito (em R$) | AC = quanto foi gasto | "
                  "SPI < 1 atrasado | CPI < 1 acima do custo | EAC = custo total projetado | VAC = desvio final previsto",
       tamanho=9, cor_fonte=CINZA, borda=False)
ws.merge_cells(start_row=gl, start_column=2, end_row=gl, end_column=18)

# ---- tabela mensal para a Curva S (alimenta o grafico do Dashboard)
CS = gl + 2
celula(ws, CS, 2, "CURVA S — VALORES ACUMULADOS (R$ mil)", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO)
ws.merge_cells(start_row=CS, start_column=2, end_row=CS, end_column=11)
meses = ["Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out"]
pv_acum = [35, 120, 245, 390, 540, 660, 780, 850]
ev_acum = [33, 112, 228, 352, None, None, None, None]
ac_acum = [36, 121, 248, 374, None, None, None, None]
cabecalho(ws, CS + 1, ["", "Mês"] + meses, cor=AZUL)
for nome_linha, dados, rr in [("PV — Valor Planejado", pv_acum, CS + 2),
                              ("EV — Valor Agregado", ev_acum, CS + 3),
                              ("AC — Custo Real", ac_acum, CS + 4)]:
    celula(ws, rr, 2, nome_linha, bold=True, cor_fundo=AZUL_CLARO)
    for j, v in enumerate(dados):
        celula(ws, rr, 3 + j, v, fmt="#,##0", alinh="center")
CS_DADOS = (CS + 1, CS + 4, 2, 2 + len(meses))  # linhas/colunas da tabela

# =============================================================== RISCOS
ws = wb.create_sheet("⚠️ Riscos")
ws.sheet_properties.tabColor = VERMELHO
ws.merge_cells("A1:O1")
estilo_titulo(ws, "A1", "REGISTRO DE RISCOS — ANÁLISE QUALITATIVA (PROBABILIDADE × IMPACTO)")
ws.row_dimensions[1].height = 30
cols = ["ID", "Data", "Categoria", "Descrição do Risco (causa → evento → efeito)", "Prob.\n(1-5)",
        "Imp.\n(1-5)", "Score\n(auto)", "Nível\n(auto)", "Estratégia", "Plano de Resposta",
        "Responsável", "Status", "Custo da\nResposta", "Próxima\nRevisão", "Exposição\n(R$, auto)"]
cabecalho(ws, 3, cols, altura=40)
larguras(ws, [6, 11, 15, 52, 6, 6, 7, 12, 12, 42, 15, 14, 12, 11, 12])

riscos = [
    ("R-01", dt.date(2026, 3, 20), "Técnico", "Qualidade dos dados legados abaixo do esperado pode atrasar a migração e o go-live", 4, 5,
     "Mitigar", "Auditoria de dados antecipada; 3 cargas de teste; equipe extra de limpeza contratada", "Roberto Lima", "Em Tratamento", 25000, dt.date(2026, 6, 26)),
    ("R-02", dt.date(2026, 3, 20), "Organizacional", "Indisponibilidade dos key users nas fases de parametrização e testes", 4, 4,
     "Mitigar", "Acordo formal de 50% de alocação assinado pelos diretores; agenda bloqueada com 4 semanas de antecedência", "Ana B. Souza", "Em Tratamento", 0, dt.date(2026, 6, 19)),
    ("R-03", dt.date(2026, 3, 25), "Comercial", "Atraso do fornecedor na entrega das customizações contratadas", 3, 4,
     "Transferir", "Multa contratual por atraso; marcos de pagamento vinculados a entregas aceitas", "Ana B. Souza", "Aberto", 0, dt.date(2026, 7, 3)),
    ("R-04", dt.date(2026, 4, 2), "Técnico", "Integração com o sistema fiscal legado mais complexa que o estimado", 3, 3,
     "Mitigar", "Prova de conceito da integração antes do desenvolvimento; consultor especialista sob demanda", "Diego Martins", "Em Tratamento", 18000, dt.date(2026, 6, 30)),
    ("R-05", dt.date(2026, 4, 10), "Externo", "Mudança na legislação fiscal exigindo reparametrização", 2, 4,
     "Aceitar", "Monitorar publicações; reserva de contingência cobre até 200 h de ajustes", "Fernanda Costa", "Aberto", 0, dt.date(2026, 7, 31)),
    ("R-06", dt.date(2026, 4, 15), "Organizacional", "Resistência dos usuários à mudança reduz adoção pós go-live", 3, 3,
     "Mitigar", "Plano de gestão da mudança; champions por área; comunicação quinzenal", "Fernanda Costa", "Em Tratamento", 12000, dt.date(2026, 7, 10)),
    ("R-07", dt.date(2026, 5, 5), "Gerenciamento", "Estouro de horas de customização além das 400 h contratadas", 3, 4,
     "Mitigar", "Comitê de mudanças rigoroso; backlog priorizado; relatório semanal de consumo de horas", "Ana B. Souza", "Em Tratamento", 0, dt.date(2026, 6, 22)),
    ("R-08", dt.date(2026, 5, 12), "Técnico", "Performance inadequada do ambiente cloud em picos de carga", 2, 3,
     "Mitigar", "Teste de carga antes do go-live; autoscaling configurado", "Roberto Lima", "Aberto", 8000, dt.date(2026, 8, 14)),
    ("R-09", dt.date(2026, 5, 20), "Externo", "Indisponibilidade do consultor sênior do fornecedor (pessoa-chave)", 2, 4,
     "Transferir", "Cláusula de substituição em 5 dias úteis com profissional equivalente", "Ana B. Souza", "Aberto", 0, dt.date(2026, 7, 17)),
    ("R-10", dt.date(2026, 6, 1), "Regulatório", "Auditoria externa durante o go-live consumindo a equipe de TI", 1, 3,
     "Aceitar", "Plano de contingência com janela alternativa de go-live em 23/10", "Ana B. Souza", "Aberto", 0, dt.date(2026, 9, 4)),
]
R0 = 4
for i, (rid, data, cat, desc, p, imp, estr, plano, resp, st, custo, rev) in enumerate(riscos):
    r = R0 + i
    celula(ws, r, 1, rid, alinh="center", bold=True)
    celula(ws, r, 2, data, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 3, cat, alinh="center")
    celula(ws, r, 4, desc, wrap=True)
    celula(ws, r, 5, p, alinh="center")
    celula(ws, r, 6, imp, alinh="center")
    celula(ws, r, 7, f"=IF(OR(E{r}=\"\",F{r}=\"\"),\"\",E{r}*F{r})", alinh="center", bold=True)
    celula(ws, r, 8,
           f"=IF(G{r}=\"\",\"\",IF(G{r}>=20,\"🔴 Crítico\",IF(G{r}>=12,\"🟠 Alto\",IF(G{r}>=6,\"🟡 Moderado\",\"🟢 Baixo\"))))",
           alinh="center")
    celula(ws, r, 9, estr, alinh="center")
    celula(ws, r, 10, plano, wrap=True)
    celula(ws, r, 11, resp, alinh="center")
    celula(ws, r, 12, st, alinh="center")
    celula(ws, r, 13, custo, fmt=FMT_MOEDA_K)
    celula(ws, r, 14, rev, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 15, f"=IF(OR(E{r}=\"\",F{r}=\"\"),\"\",E{r}/5*F{r}/5*'📋 TAP'!$C$10)", fmt=FMT_MOEDA_K)
RF = R0 + len(riscos) - 1
RMAX = RF + 20
for r in range(RF + 1, RMAX + 1):
    celula(ws, r, 7, f"=IF(OR(E{r}=\"\",F{r}=\"\"),\"\",E{r}*F{r})", alinh="center", bold=True)
    celula(ws, r, 8,
           f"=IF(G{r}=\"\",\"\",IF(G{r}>=20,\"🔴 Crítico\",IF(G{r}>=12,\"🟠 Alto\",IF(G{r}>=6,\"🟡 Moderado\",\"🟢 Baixo\"))))",
           alinh="center")
    for cc in list(range(1, 7)) + list(range(9, 15)):
        ws.cell(row=r, column=cc).border = BORDA
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:O{RF}"
dropdown(ws, "ListaCatRisco", f"C{R0}:C{RMAX}")
dropdown(ws, "Lista15", f"E{R0}:F{RMAX}")
dropdown(ws, "ListaEstrategia", f"I{R0}:I{RMAX}")
dropdown(ws, "ListaStatusRisco", f"L{R0}:L{RMAX}")

for txt, cor in [("Crítico", VERMELHO_CLARO), ("Alto", "FBE2D5"), ("Moderado", AMARELO_CLARO), ("Baixo", VERDE_CLARO)]:
    ws.conditional_formatting.add(f"H{R0}:H{RMAX}", FormulaRule(
        formula=[f'NOT(ISERROR(SEARCH("{txt}",H{R0})))'], fill=fill(cor), stopIfTrue=True))
ws.conditional_formatting.add(f"G{R0}:G{RMAX}", ColorScaleRule(
    start_type="num", start_value=1, start_color="63BE7B",
    mid_type="num", mid_value=12, mid_color="FFEB84",
    end_type="num", end_value=25, end_color="F8696B"))

# ---- matriz 5x5 (a direita da tabela)
MC = 17  # coluna Q
celula(ws, 3, MC, "MATRIZ DE RISCOS 5×5 (contagem automática)", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
ws.merge_cells(start_row=3, start_column=MC, end_row=3, end_column=MC + 6)
celula(ws, 4, MC, "PROBABILIDADE ↓ / IMPACTO →", bold=True, tamanho=8, alinh="center", cor_fundo=CINZA_CLARO)
for j in range(1, 6):
    celula(ws, 4, MC + j, j, bold=True, alinh="center", cor_fundo=CINZA_CLARO)
cores_matriz = {  # score -> cor
    (1, 4): VERDE_CLARO, (5, 9): AMARELO_CLARO, (10, 14): "FBE2D5", (15, 25): VERMELHO_CLARO,
}
for p in range(5, 0, -1):
    rr = 4 + (6 - p)
    celula(ws, rr, MC, p, bold=True, alinh="center", cor_fundo=CINZA_CLARO)
    for imp in range(1, 6):
        score = p * imp
        cor = next(c for (a, b), c in cores_matriz.items() if a <= score <= b)
        celula(ws, rr, MC + imp,
               f"=COUNTIFS($E${R0}:$E${RMAX},{p},$F${R0}:$F${RMAX},{imp},$L${R0}:$L${RMAX},\"<>Encerrado\")",
               alinh="center", bold=True, cor_fundo=cor)
ws.column_dimensions[get_column_letter(MC)].width = 14
for j in range(1, 6):
    ws.column_dimensions[get_column_letter(MC + j)].width = 6

RISCO_NIVEL = f"'⚠️ Riscos'!$H${R0}:$H${RMAX}"
RISCO_STATUS = f"'⚠️ Riscos'!$L${R0}:$L${RMAX}"

# ========================================================== STAKEHOLDERS
ws = wb.create_sheet("👥 Stakeholders")
ws.sheet_properties.tabColor = LARANJA
ws.merge_cells("A1:L1")
estilo_titulo(ws, "A1", "REGISTRO DE PARTES INTERESSADAS — MATRIZ PODER × INTERESSE")
ws.row_dimensions[1].height = 30
cols = ["ID", "Nome", "Cargo / Organização", "Papel no Projeto", "Poder\n(1-5)", "Interesse\n(1-5)",
        "Quadrante (auto)", "Engajamento\nAtual", "Engajamento\nDesejado", "Estratégia de Engajamento",
        "Canal / Frequência", "Responsável"]
cabecalho(ws, 3, cols, altura=40)
larguras(ws, [5, 20, 24, 20, 8, 9, 18, 13, 13, 42, 20, 14])
stakeholders = [
    ("Carlos Mendes", "Diretor de Operações", "Patrocinador", 5, 5, "Apoiador", "Líder",
     "Reunião executiva quinzenal; decisões de escopo e verba passam por ele", "Reunião / quinzenal", "Ana B. Souza"),
    ("Marta Oliveira", "CFO", "Comitê Diretivo", 5, 3, "Neutro", "Apoiador",
     "Apresentar ROI e ganhos no fechamento contábil; envolver na aprovação do orçamento", "Status report / mensal", "Ana B. Souza"),
    ("Paulo Ferreira", "Gerente de TI", "Líder técnico interno", 4, 5, "Apoiador", "Apoiador",
     "Envolver nas decisões de arquitetura e infraestrutura", "Daily técnica / semanal", "Roberto Lima"),
    ("Key users (12)", "Coordenadores das áreas", "Validação dos processos", 3, 5, "Resistente", "Apoiador",
     "Workshops de cocriação; reconhecer esforço junto às diretorias; plano de gestão da mudança", "Workshop / semanal", "Fernanda Costa"),
    ("Usuários finais (320)", "Operação", "Usuários do sistema", 1, 4, "Desinformado", "Neutro",
     "Comunicação quinzenal; champions por área; treinamento próximo ao go-live", "E-mail+intranet / quinzenal", "Fernanda Costa"),
    ("TechSolutions", "Fornecedor do ERP", "Implantadora", 4, 4, "Apoiador", "Apoiador",
     "Governança contratual; reunião de steering mensal; SLAs monitorados", "Steering / mensal", "Ana B. Souza"),
    ("Auditoria Interna", "Compliance", "Controles e conformidade", 3, 2, "Neutro", "Neutro",
     "Informar marcos e evidências de controles; convidar para o UAT", "Relatório / por marco", "Ana B. Souza"),
    ("Sindicato/RH", "Recursos Humanos", "Impactos na operação", 2, 2, "Desinformado", "Neutro",
     "Comunicado formal sobre mudanças de rotina; canal de dúvidas", "Comunicado / por fase", "Fernanda Costa"),
]
S0 = 4
for i, (nome, cargo, papel, poder, inter, eng_a, eng_d, estrat, canal, resp) in enumerate(stakeholders):
    r = S0 + i
    celula(ws, r, 1, f"S-{i+1:02d}", alinh="center", bold=True)
    celula(ws, r, 2, nome)
    celula(ws, r, 3, cargo)
    celula(ws, r, 4, papel)
    celula(ws, r, 5, poder, alinh="center")
    celula(ws, r, 6, inter, alinh="center")
    celula(ws, r, 7,
           f"=IF(OR(E{r}=\"\",F{r}=\"\"),\"\",IF(AND(E{r}>=3,F{r}>=3),\"🔴 Gerenciar de Perto\","
           f"IF(E{r}>=3,\"🟠 Manter Satisfeito\",IF(F{r}>=3,\"🟡 Manter Informado\",\"🟢 Monitorar\"))))",
           alinh="center")
    celula(ws, r, 8, eng_a, alinh="center")
    celula(ws, r, 9, eng_d, alinh="center")
    celula(ws, r, 10, estrat, wrap=True)
    celula(ws, r, 11, canal, alinh="center")
    celula(ws, r, 12, resp, alinh="center")
SF = S0 + len(stakeholders) - 1
SMAX = SF + 15
for r in range(SF + 1, SMAX + 1):
    celula(ws, r, 7,
           f"=IF(OR(E{r}=\"\",F{r}=\"\"),\"\",IF(AND(E{r}>=3,F{r}>=3),\"🔴 Gerenciar de Perto\","
           f"IF(E{r}>=3,\"🟠 Manter Satisfeito\",IF(F{r}>=3,\"🟡 Manter Informado\",\"🟢 Monitorar\"))))",
           alinh="center")
    for cc in list(range(1, 7)) + list(range(8, 13)):
        ws.cell(row=r, column=cc).border = BORDA
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:L{SF}"
dropdown(ws, "Lista15", f"E{S0}:F{SMAX}")
dropdown(ws, "ListaEngajamento", f"H{S0}:I{SMAX}")
# destaca quando o engajamento atual difere do desejado (exige acao do GP)
ws.conditional_formatting.add(f"H{S0}:H{SMAX}", FormulaRule(
    formula=[f'AND($H{S0}<>"",$I{S0}<>"",$H{S0}<>$I{S0})'], fill=fill(AMARELO_CLARO)))
for txt, cor in [("Gerenciar de Perto", VERMELHO_CLARO), ("Manter Satisfeito", "FBE2D5"),
                 ("Manter Informado", AMARELO_CLARO), ("Monitorar", VERDE_CLARO)]:
    ws.conditional_formatting.add(f"G{S0}:G{SMAX}", FormulaRule(
        formula=[f'NOT(ISERROR(SEARCH("{txt}",G{S0})))'], fill=fill(cor), stopIfTrue=True))

# ============================================================= MUDANÇAS
ws = wb.create_sheet("🔄 Mudanças")
ws.sheet_properties.tabColor = AZUL
ws.merge_cells("A1:L1")
estilo_titulo(ws, "A1", "CONTROLE INTEGRADO DE MUDANÇAS (CHANGE LOG)")
ws.row_dimensions[1].height = 30
cols = ["ID", "Data", "Solicitante", "Descrição da Mudança", "Justificativa", "Impacto no\nEscopo",
        "Impacto no\nPrazo (dias)", "Impacto no\nCusto (R$)", "Prioridade", "Status", "Decisão /\nData", "Aprovador"]
cabecalho(ws, 3, cols, altura=40)
larguras(ws, [7, 11, 16, 42, 34, 14, 11, 13, 11, 14, 12, 16])
mudancas = [
    ("MUD-01", dt.date(2026, 4, 22), "Diretoria Comercial", "Incluir relatório de comissões customizado no módulo Vendas",
     "Relatório atual é feito manualmente e consome 40 h/mês", "Aumenta", 5, 22000, "🟡 Média", "Aprovada", dt.date(2026, 5, 4), "Comitê (CCB)"),
    ("MUD-02", dt.date(2026, 5, 15), "Key user Financeiro", "Alterar fluxo de aprovação de pagamentos para 3 níveis",
     "Política de alçadas revisada pela CFO", "Neutro", 3, 8500, "🔴 Alta", "Aprovada", dt.date(2026, 5, 22), "Comitê (CCB)"),
    ("MUD-03", dt.date(2026, 5, 28), "Gerente de TI", "Antecipar integração com e-commerce (fase 2) para o go-live",
     "Oportunidade comercial de fim de ano", "Aumenta", 25, 95000, "🟡 Média", "Rejeitada", dt.date(2026, 6, 5), "Patrocinador"),
    ("MUD-04", dt.date(2026, 6, 8), "Fornecedor", "Substituir ferramenta de carga de dados por versão mais recente",
     "Versão atual será descontinuada pelo fabricante", "Neutro", 0, 0, "🟢 Baixa", "Em Análise", None, ""),
]
M0 = 4
for i, vals in enumerate(mudancas):
    r = M0 + i
    for j, v in enumerate(vals, start=1):
        fmt = FMT_DATA if j in (2, 11) else (FMT_MOEDA_K if j == 8 else None)
        celula(ws, r, j, v, fmt=fmt, alinh="center" if j not in (4, 5) else "left", wrap=j in (4, 5))
MF = M0 + len(mudancas) - 1
MMAX = MF + 15
for r in range(MF + 1, MMAX + 1):
    for cc in range(1, 13):
        ws.cell(row=r, column=cc).border = BORDA
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:L{MF}"
dropdown(ws, "ListaPrioridade", f"I{M0}:I{MMAX}")
dropdown(ws, "ListaStatusMudanca", f"J{M0}:J{MMAX}")
for txt, cor in [("Aprovada", VERDE_CLARO), ("Rejeitada", VERMELHO_CLARO), ("Em Análise", AMARELO_CLARO),
                 ("Implementada", AZUL_CLARO)]:
    ws.conditional_formatting.add(f"J{M0}:J{MMAX}", FormulaRule(
        formula=[f'NOT(ISERROR(SEARCH("{txt}",J{M0})))'], fill=fill(cor), stopIfTrue=True))
MUD_STATUS = f"'🔄 Mudanças'!$J${M0}:$J${MMAX}"
MUD_CUSTO = f"'🔄 Mudanças'!$H${M0}:$H${MMAX}"

# ============================================================== QUESTÕES
ws = wb.create_sheet("🚧 Questões")
ws.sheet_properties.tabColor = LARANJA
ws.merge_cells("A1:K1")
estilo_titulo(ws, "A1", "REGISTRO DE QUESTÕES E IMPEDIMENTOS (ISSUE LOG)")
ws.row_dimensions[1].height = 30
cols = ["ID", "Data\nAbertura", "Descrição da Questão", "Origem /\nRelacionada a", "Prioridade",
        "Responsável", "Prazo", "Status", "Dias em\nAberto (auto)", "Alerta\n(auto)", "Resolução"]
cabecalho(ws, 3, cols, altura=40)
larguras(ws, [7, 11, 48, 14, 11, 16, 11, 14, 10, 8, 40])
questoes = [
    ("Q-01", dt.date(2026, 5, 18), "De-para de plano de contas com 340 contas sem correspondência", "R-01",
     "🔴 Alta", "Roberto Lima", dt.date(2026, 6, 10), "Em Tratamento", "Workshop com contabilidade agendado; 60% mapeado"),
    ("Q-02", dt.date(2026, 5, 25), "Key user de Suprimentos alocado em outro projeto da diretoria", "R-02",
     "🔴 Alta", "Ana B. Souza", dt.date(2026, 6, 5), "Escalada", "Escalado ao patrocinador em 01/06"),
    ("Q-03", dt.date(2026, 6, 1), "Ambiente de QA instável após atualização do fornecedor", "1.2.2",
     "🟡 Média", "Roberto Lima", dt.date(2026, 6, 12), "Em Tratamento", "Chamado aberto no fornecedor (SLA 5 dias)"),
    ("Q-04", dt.date(2026, 6, 3), "Definição pendente da política de alçadas para o fluxo de compras", "1.3.2",
     "🔴 Alta", "Diego Martins", dt.date(2026, 6, 19), "Aberta", ""),
    ("Q-05", dt.date(2026, 4, 14), "Acesso VPN dos consultores externos bloqueado pela segurança", "1.2",
     "🟡 Média", "Paulo Ferreira", dt.date(2026, 4, 24), "Resolvida", "Política de acesso temporário aprovada em 22/04"),
    ("Q-06", dt.date(2026, 6, 9), "Sala de treinamento indisponível nas datas planejadas de setembro", "1.6",
     "🟢 Baixa", "Fernanda Costa", dt.date(2026, 7, 15), "Aberta", ""),
]
Q0 = 4
for i, (qid, data, desc, orig, prio, resp, prazo, st, resol) in enumerate(questoes):
    r = Q0 + i
    celula(ws, r, 1, qid, alinh="center", bold=True)
    celula(ws, r, 2, data, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 3, desc, wrap=True)
    celula(ws, r, 4, orig, alinh="center")
    celula(ws, r, 5, prio, alinh="center")
    celula(ws, r, 6, resp, alinh="center")
    celula(ws, r, 7, prazo, fmt=FMT_DATA, alinh="center")
    celula(ws, r, 8, st, alinh="center")
    celula(ws, r, 9, f"=IF(B{r}=\"\",\"\",IF(H{r}=\"Resolvida\",\"—\",TODAY()-B{r}))", alinh="center")
    celula(ws, r, 10, f"=IF(OR(G{r}=\"\",H{r}=\"Resolvida\"),\"\",IF(TODAY()>G{r},\"🔴\",IF(G{r}-TODAY()<=5,\"🟡\",\"🟢\")))", alinh="center")
    celula(ws, r, 11, resol, wrap=True)
QF = Q0 + len(questoes) - 1
QMAX = QF + 20
for r in range(QF + 1, QMAX + 1):
    celula(ws, r, 9, f"=IF(B{r}=\"\",\"\",IF(H{r}=\"Resolvida\",\"—\",TODAY()-B{r}))", alinh="center")
    celula(ws, r, 10, f"=IF(OR(G{r}=\"\",H{r}=\"Resolvida\"),\"\",IF(TODAY()>G{r},\"🔴\",IF(G{r}-TODAY()<=5,\"🟡\",\"🟢\")))", alinh="center")
    for cc in list(range(1, 9)) + [11]:
        ws.cell(row=r, column=cc).border = BORDA
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:K{QF}"
dropdown(ws, "ListaPrioridade", f"E{Q0}:E{QMAX}")
dropdown(ws, "ListaStatusQuestao", f"H{Q0}:H{QMAX}")
for txt, cor in [("Resolvida", VERDE_CLARO), ("Escalada", VERMELHO_CLARO), ("Em Tratamento", AMARELO_CLARO)]:
    ws.conditional_formatting.add(f"H{Q0}:H{QMAX}", FormulaRule(
        formula=[f'NOT(ISERROR(SEARCH("{txt}",H{Q0})))'], fill=fill(cor), stopIfTrue=True))
QUEST_STATUS = f"'🚧 Questões'!$H${Q0}:$H${QMAX}"

# ================================================================ LIÇÕES
ws = wb.create_sheet("📚 Lições")
ws.sheet_properties.tabColor = VERDE
ws.merge_cells("A1:I1")
estilo_titulo(ws, "A1", "LIÇÕES APRENDIDAS — REGISTRE DURANTE O PROJETO, NÃO SÓ NO FIM")
ws.row_dimensions[1].height = 30
cols = ["ID", "Data", "Fase", "Categoria", "O que aconteceu", "O que funcionou / não funcionou",
        "Recomendação para projetos futuros", "Impacto", "Autor"]
cabecalho(ws, 3, cols, altura=30)
larguras(ws, [7, 11, 15, 14, 40, 40, 40, 11, 16])
licoes = [
    ("LA-01", dt.date(2026, 4, 15), "Planejamento", "Pessoas", "Workshops de levantamento com todas as áreas juntas geravam dispersão",
     "Sessões separadas por área, com pauta fechada, foram 2× mais produtivas",
     "Planejar workshops por área desde o início, com no máximo 8 participantes", "🟢 Positivo", "Fernanda Costa"),
    ("LA-02", dt.date(2026, 5, 10), "Execução", "Fornecedores", "Contrato sem critério objetivo de 'entrega aceita' gerou disputa no 1º pagamento",
     "Não funcionou: aceite implícito. Funcionou: addendum com checklist de aceitação",
     "Definir critérios de aceitação mensuráveis em contrato, antes da assinatura", "🔴 Negativo", "Ana B. Souza"),
    ("LA-03", dt.date(2026, 5, 28), "Execução", "Tecnologia", "Carga de teste antecipada revelou problemas de dados 3 meses antes do previsto",
     "Antecipar a 1ª carga de teste deu tempo hábil para corrigir a qualidade dos dados",
     "Sempre executar uma carga de dados exploratória na primeira metade do projeto", "🟢 Positivo", "Roberto Lima"),
    ("LA-04", dt.date(2026, 6, 5), "Execução", "Comunicação", "Boletim quinzenal por e-mail tinha baixa leitura (18%)",
     "Mural na intranet + 5 min nas reuniões de equipe das áreas funcionou melhor",
     "Validar o canal de comunicação com uma amostra de usuários antes de padronizar", "🟡 Neutro", "Fernanda Costa"),
]
LI0 = 4
for i, vals in enumerate(licoes):
    r = LI0 + i
    for j, v in enumerate(vals, start=1):
        celula(ws, r, j, v, fmt=FMT_DATA if j == 2 else None,
               alinh="center" if j in (1, 2, 3, 4, 8, 9) else "left", wrap=j in (5, 6, 7))
LIF = LI0 + len(licoes) - 1
for r in range(LIF + 1, LIF + 16):
    for cc in range(1, 10):
        ws.cell(row=r, column=cc).border = BORDA
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:I{LIF}"
dropdown(ws, "ListaFase", f"C{LI0}:C{LIF + 15}")
dropdown(ws, "ListaCatLicao", f"D{LI0}:D{LIF + 15}")

# ============================================================= DASHBOARD
ws = wb.create_sheet("📊 Dashboard", 1)
ws.sheet_properties.tabColor = AZUL_ESCURO
ws.sheet_view.showGridLines = False
larguras(ws, [2, 16, 14, 14, 14, 2, 16, 14, 14, 14, 2, 16, 13, 13, 13, 2])

ws.merge_cells("B2:O2")
estilo_titulo(ws, "B2", "DASHBOARD EXECUTIVO DO PROJETO", tamanho=20)
ws.row_dimensions[2].height = 40
ws.merge_cells("B3:O3")
c = ws["B3"]
# DAY/MONTH/YEAR em vez de TEXT() para nao depender do idioma do Excel
c.value = ("='📋 TAP'!C4&\"  |  GP: \"&'📋 TAP'!C6&\"  |  Atualizado em: \""
           "&DAY(TODAY())&\"/\"&MONTH(TODAY())&\"/\"&YEAR(TODAY())")
c.font = Font(size=11, italic=True, color=CINZA)
c.alignment = Alignment(horizontal="center")

# ---------- linha de KPIs (cartoes)
def cartao(col_ini, titulo, formula, fmt=None, cor=AZUL):
    ws.merge_cells(start_row=5, start_column=col_ini, end_row=5, end_column=col_ini + 1)
    celula(ws, 5, col_ini, titulo, bold=True, cor_fundo=cor, cor_fonte=BRANCO, alinh="center", tamanho=10)
    ws.cell(row=5, column=col_ini + 1).fill = fill(cor)
    ws.cell(row=5, column=col_ini + 1).border = BORDA
    ws.merge_cells(start_row=6, start_column=col_ini, end_row=7, end_column=col_ini + 1)
    c = celula(ws, 6, col_ini, formula, fmt=fmt, bold=True, alinh="center", tamanho=22, cor_fundo=CINZA_CLARO)
    for rr in (6, 7):
        for cc in (col_ini, col_ini + 1):
            ws.cell(row=rr, column=cc).border = BORDA
            if not (rr == 6 and cc == col_ini):
                ws.cell(row=rr, column=cc).fill = fill(CINZA_CLARO)
    return c

ws.row_dimensions[6].height = 26
ws.row_dimensions[7].height = 16
cartao(2, "SAÚDE GERAL",
       f"=IF(AND('💰 EVM'!N{TR}>=0.95,'💰 EVM'!M{TR}>=0.95),\"🟢 NO RUMO\","
       f"IF(OR('💰 EVM'!N{TR}<0.85,'💰 EVM'!M{TR}<0.85),\"🔴 CRÍTICO\",\"🟡 ATENÇÃO\"))")
cartao(4, "AVANÇO FÍSICO", f"=AVERAGE({CRONO_PCT})", fmt=FMT_PCT)
cartao(7, "SPI (PRAZO)", f"='💰 EVM'!M{TR}", fmt=FMT_NUM2)
cartao(9, "CPI (CUSTO)", f"='💰 EVM'!N{TR}", fmt=FMT_NUM2)
cartao(12, "DIAS P/ TÉRMINO", "='📋 TAP'!C9-TODAY()", fmt="0")
cartao(14, "RISCOS CRÍT.+ALTOS",
       f"=SUMPRODUCT(--(ISNUMBER(SEARCH(\"Crítico\",{RISCO_NIVEL}))))+SUMPRODUCT(--(ISNUMBER(SEARCH(\"Alto\",{RISCO_NIVEL}))))",
       fmt="0", cor=VERMELHO)
# semaforo nos KPIs SPI/CPI
for cel in ("G6", "I6"):
    ws.conditional_formatting.add(cel, CellIsRule(operator="lessThan", formula=["0.85"], fill=fill(VERMELHO_CLARO)))
    ws.conditional_formatting.add(cel, CellIsRule(operator="between", formula=["0.85", "0.9499"], fill=fill(AMARELO_CLARO)))
    ws.conditional_formatting.add(cel, CellIsRule(operator="greaterThanOrEqual", formula=["0.95"], fill=fill(VERDE_CLARO)))

# ---------- bloco TAREFAS
r = 9
celula(ws, r, 2, "TAREFAS DO CRONOGRAMA", bold=True, cor_fundo=AZUL_ESCURO, cor_fonte=BRANCO, alinh="center")
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
for cc in range(3, 6):
    ws.cell(row=r, column=cc).border = BORDA
tarefas_kpi = [
    ("✅ Concluídas", f"=COUNTIF({CRONO_STATUS},\"*Concluída*\")", VERDE_CLARO),
    ("🟡 Em Andamento", f"=COUNTIF({CRONO_STATUS},\"*Em Andamento*\")", AMARELO_CLARO),
    ("🔴 Atrasadas", f"=COUNTIF({CRONO_STATUS},\"*Atrasada*\")", VERMELHO_CLARO),
    ("⚪ Não Iniciadas", f"=COUNTIF({CRONO_STATUS},\"*Não Iniciada*\")", CINZA_CLARO),
]
for i, (rot, f, cor) in enumerate(tarefas_kpi):
    rr = r + 1 + i
    celula(ws, rr, 2, rot, bold=True, cor_fundo=cor)
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    celula(ws, rr, 3, f, bold=True, alinh="center")
    ws.cell(row=rr, column=4).border = BORDA
    celula(ws, rr, 5, f"=IFERROR(C{rr}/SUM($C${r+1}:$C${r+4}),0)", fmt=FMT_PCT, alinh="center")

# ---------- bloco RISCOS
celula(ws, r, 7, "RISCOS ABERTOS POR NÍVEL", bold=True, cor_fundo=VERMELHO, cor_fonte=BRANCO, alinh="center")
ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=10)
for cc in range(8, 11):
    ws.cell(row=r, column=cc).border = BORDA
riscos_kpi = [
    ("🔴 Crítico", "Crítico", VERMELHO_CLARO),
    ("🟠 Alto", "Alto", "FBE2D5"),
    ("🟡 Moderado", "Moderado", AMARELO_CLARO),
    ("🟢 Baixo", "Baixo", VERDE_CLARO),
]
for i, (rot, chave, cor) in enumerate(riscos_kpi):
    rr = r + 1 + i
    celula(ws, rr, 7, rot, bold=True, cor_fundo=cor)
    ws.merge_cells(start_row=rr, start_column=8, end_row=rr, end_column=9)
    celula(ws, rr, 8,
           f"=SUMPRODUCT(--(ISNUMBER(SEARCH(\"{chave}\",{RISCO_NIVEL}))),--({RISCO_STATUS}<>\"Encerrado\"))",
           bold=True, alinh="center")
    ws.cell(row=rr, column=9).border = BORDA
    celula(ws, rr, 10, "", borda=True)

# ---------- bloco PENDÊNCIAS
celula(ws, r, 12, "PENDÊNCIAS E MUDANÇAS", bold=True, cor_fundo=LARANJA, cor_fonte=BRANCO, alinh="center")
ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=15)
for cc in range(13, 16):
    ws.cell(row=r, column=cc).border = BORDA
pend = [
    ("Questões abertas", f"=SUMPRODUCT(--({QUEST_STATUS}<>\"\"),--({QUEST_STATUS}<>\"Resolvida\"))"),
    ("Questões escaladas", f"=COUNTIF({QUEST_STATUS},\"Escalada\")"),
    ("Mudanças em análise", f"=COUNTIF({MUD_STATUS},\"Em Análise\")"),
    ("Custo de mudanças aprov.", f"=SUMIFS({MUD_CUSTO},{MUD_STATUS},\"Aprovada\")"),
]
for i, (rot, f) in enumerate(pend):
    rr = r + 1 + i
    ws.merge_cells(start_row=rr, start_column=12, end_row=rr, end_column=13)
    celula(ws, rr, 12, rot, bold=True, cor_fundo=CINZA_CLARO)
    ws.cell(row=rr, column=13).border = BORDA
    ws.merge_cells(start_row=rr, start_column=14, end_row=rr, end_column=15)
    celula(ws, rr, 14, f, bold=True, alinh="center", fmt=FMT_MOEDA_K if i == 3 else "0")
    ws.cell(row=rr, column=15).border = BORDA

# ---------- bloco EVM resumo
r = 15
celula(ws, r, 2, "RESUMO FINANCEIRO (EVM)", bold=True, cor_fundo=AMARELO, alinh="center")
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
for cc in range(3, 6):
    ws.cell(row=r, column=cc).border = BORDA
evm_kpi = [
    ("Orçamento (BAC)", f"='💰 EVM'!E{TR}"),
    ("Valor Agregado (EV)", f"='💰 EVM'!I{TR}"),
    ("Custo Real (AC)", f"='💰 EVM'!J{TR}"),
    ("Projeção Final (EAC)", f"='💰 EVM'!O{TR}"),
    ("Desvio Final (VAC)", f"='💰 EVM'!Q{TR}"),
]
for i, (rot, f) in enumerate(evm_kpi):
    rr = r + 1 + i
    celula(ws, rr, 2, rot, bold=True, cor_fundo=CINZA_CLARO)
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=5)
    celula(ws, rr, 3, f, fmt=FMT_MOEDA_K, bold=True, alinh="center")
    for cc in (4, 5):
        ws.cell(row=rr, column=cc).border = BORDA
ws.conditional_formatting.add(f"C{r+5}", CellIsRule(operator="lessThan", formula=["0"], fill=fill(VERMELHO_CLARO)))
ws.conditional_formatting.add(f"C{r+5}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=fill(VERDE_CLARO)))

# ---------- proximos marcos
celula(ws, r, 7, "PRÓXIMOS MARCOS", bold=True, cor_fundo=AZUL, cor_fonte=BRANCO, alinh="center")
ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=10)
for cc in range(8, 11):
    ws.cell(row=r, column=cc).border = BORDA
marcos_dash = [(12, ), (14, ), (19, ), (21, )]  # IDs no cronograma
for i, (tid, ) in enumerate(marcos_dash):
    rr = r + 1 + i
    lin = L0 + tid - 1
    ws.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=8)
    celula(ws, rr, 7, f"='📅 Cronograma'!C{lin}", cor_fundo=CINZA_CLARO)
    ws.cell(row=rr, column=8).border = BORDA
    celula(ws, rr, 9, f"='📅 Cronograma'!G{lin}", fmt=FMT_DATA, alinh="center")
    celula(ws, rr, 10, f"='📅 Cronograma'!J{lin}", alinh="center")

# ---------- graficos
# Curva S
lc = LineChart()
lc.title = "Curva S — PV × EV × AC (R$ mil)"
lc.style = 12
lc.height = 8.5
lc.width = 17
r1, r2, c1, c2 = CS_DADOS
dados = Reference(wb["💰 EVM"], min_row=r1 + 1, max_row=r2, min_col=c1, max_col=c2)
cats = Reference(wb["💰 EVM"], min_row=r1, min_col=c1 + 1, max_col=c2)
lc.add_data(dados, titles_from_data=True, from_rows=True)
lc.set_categories(cats)
lc.y_axis.title = "R$ mil (acum.)"
ws.add_chart(lc, "B22")

# Tarefas por status (barras)
bc = BarChart()
bc.type = "col"
bc.title = "Tarefas por Status"
bc.style = 10
bc.height = 8.5
bc.width = 8.5
dados = Reference(ws, min_row=10, max_row=13, min_col=3)
cats = Reference(ws, min_row=10, max_row=13, min_col=2)
bc.add_data(dados)
bc.set_categories(cats)
bc.legend = None
ws.add_chart(bc, "K22")

# Riscos por nivel (pizza)
pc = PieChart()
pc.title = "Riscos Abertos por Nível"
pc.height = 8.5
pc.width = 8.5
dados = Reference(ws, min_row=10, max_row=13, min_col=8)
cats = Reference(ws, min_row=10, max_row=13, min_col=7)
pc.add_data(dados)
pc.set_categories(cats)
ws.add_chart(pc, "G22")

# ordena as abas: Ajuda por ultimo nao; ordem desejada
ordem = ["📊 Dashboard", "📋 TAP", "🗂️ EAP", "📅 Cronograma", "💰 EVM", "⚠️ Riscos",
         "👥 Stakeholders", "🔄 Mudanças", "🚧 Questões", "📚 Lições", "📖 Ajuda", "Listas"]
wb._sheets = [wb[n] for n in ordem]
wb.active = 0

ARQ = "Planilha_Gerenciamento_Projetos_PMI.xlsx"
wb.save(ARQ)
print(f"OK: {ARQ} gerado com {len(wb.sheetnames)} abas: {wb.sheetnames}")
