import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

# ============================================
# AUTOMAÇÃO DE PLANILHAS EXCEL
# Desenvolvido por: Rafael Gonçalves
# Descrição: Organiza e formata planilhas automaticamente
# ============================================

def criar_relatorio_vendas(dados, nome_arquivo="relatorio_vendas.xlsx"):
    """
    Cria um relatório de vendas formatado automaticamente.
    
    dados: lista de vendas no formato [nome, produto, valor, data]
    nome_arquivo: nome do arquivo Excel a ser gerado
    """
    
    # Cria a planilha
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório de Vendas"

    # Estilo do cabeçalho
    cabecalho_font = Font(bold=True, color="FFFFFF", size=12)
    cabecalho_fill = PatternFill("solid", fgColor="2E86AB")git init
    cabecalho_align = Alignment(horizontal="center")

    # Cria o cabeçalho
    colunas = ["Cliente", "Produto", "Valor (R$)", "Data", "Status"]
    for col, titulo in enumerate(colunas, start=1):
        celula = ws.cell(row=1, column=col, value=titulo)
        celula.font = cabecalho_font
        celula.fill = cabecalho_fill
        celula.alignment = cabecalho_align

    # Preenche os dados
    total = 0
    for row, (nome, produto, valor, data) in enumerate(dados, start=2):
        ws.cell(row=row, column=1, value=nome)
        ws.cell(row=row, column=2, value=produto)
        ws.cell(row=row, column=3, value=valor)
        ws.cell(row=row, column=4, value=data)
        ws.cell(row=row, column=5, value="✓ Pago")
        total += valor

    # Linha de total
    linha_total = len(dados) + 2
    ws.cell(row=linha_total, column=2, value="TOTAL").font = Font(bold=True)
    ws.cell(row=linha_total, column=3, value=total).font = Font(bold=True, color="FF0000")

    # Ajusta largura das colunas automaticamente
    larguras = [20, 20, 15, 15, 12]
    for col, largura in enumerate(larguras, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = largura

    # Salva o arquivo
    wb.save(nome_arquivo)
    print(f"Relatório gerado com sucesso: {nome_arquivo}")
    print(f"Total de vendas: R$ {total:.2f}")


# ============================================
# EXEMPLO DE USO
# ============================================

vendas = [
    ["João Silva",    "Consultoria Python",  500.00, "26/03/2026"],
    ["Maria Souza",   "Bot WhatsApp",        350.00, "26/03/2026"],
    ["Carlos Lima",   "Automação Planilha",  400.00, "26/03/2026"],
    ["Ana Paula",     "Raspagem de Dados",   600.00, "26/03/2026"],
    ["Pedro Costa",   "Consultoria Python",  500.00, "26/03/2026"],
]

criar_relatorio_vendas(vendas)