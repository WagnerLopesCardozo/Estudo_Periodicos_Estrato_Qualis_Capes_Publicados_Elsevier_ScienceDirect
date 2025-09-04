import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


def selecionar_arquivo(titulo):
    """Abre janela para o usuário selecionar um arquivo Excel"""
    Tk().withdraw()
    caminho = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    return caminho


def gerar_pdf(relatorio, resumo, grafico_path, saida_pdf="Relatorio_ISSN.pdf"):
    """Gera relatório em PDF"""
    doc = SimpleDocTemplate(saida_pdf, pagesize=A4)
    elementos = []
    estilos = getSampleStyleSheet()

    # Título
    elementos.append(Paragraph("<b>Relatório de Comparação de Periódicos (ISSN)</b>", estilos["Title"]))
    elementos.append(Spacer(1, 12))

    # Tabela com dados detalhados
    tabela_dados = [["ISSN", "Nome do Periódico", "Estrato Qualis"]]
    tabela_dados.extend(relatorio)

    tabela = Table(tabela_dados, repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elementos.append(tabela)
    elementos.append(Spacer(1, 20))

    # Resumo por extrato
    elementos.append(Paragraph("<b>Resumo por Estrato Qualis:</b>", estilos["Heading2"]))
    resumo_tab = [["Estrato Qualis", "Quantidade"]]
    resumo_tab.extend([[k, v] for k, v in resumo.items()])

    tabela_resumo = Table(resumo_tab, repeatRows=1)
    tabela_resumo.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elementos.append(tabela_resumo)
    elementos.append(Spacer(1, 20))

    # Gráfico
    elementos.append(Paragraph("<b>Distribuição Gráfica:</b>", estilos["Heading2"]))
    elementos.append(Image(grafico_path, width=400, height=250))

    doc.build(elementos)


def main():
    print("\n=== Comparação de ISSN entre Tabelas ===")

    # Selecionar tabelas
    tabela1_path = selecionar_arquivo("Selecione a Tabela Estrato Qualis CAPES")
    tabela2_path = selecionar_arquivo("Selecione a Tabela Catálogo Completo Elsevier - Science Direct")

    # Carregar tabelas
    tabela1 = pd.read_excel(tabela1_path)
    tabela2 = pd.read_excel(tabela2_path)

    # Padronizar nomes de colunas
    tabela1.columns = tabela1.columns.str.strip()
    tabela2.columns = tabela2.columns.str.strip()

    # Extrair ISSN
    issn_capes = tabela1.iloc[:, 0].astype(str).str.strip()   # primeira coluna (aqui selecionamos a coluna que contém a informação ISSN)
    issn_elsevier = tabela2.iloc[:, 2].astype(str).str.strip()  # terceira coluna (o mesmo aqui, onde selecionamos a coluna que contém a informação ISSN)

    # Encontrar ISSN comuns
    comuns = set(issn_capes).intersection(set(issn_elsevier))

    relatorio = []
    resumo = {k: 0 for k in ["A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4", "C"]}

    print("\n=== Periódicos Encontrados ===")
    for _, row in tabela1.iterrows():
        if str(row.iloc[0]).strip() in comuns:
            issn = str(row.iloc[0]).strip()
            nome = row.iloc[1]  # Nome do periódico
            qualis = str(row.iloc[2]).strip()  # Estrato Qualis
            relatorio.append([issn, nome, qualis])
            if qualis in resumo:
                resumo[qualis] += 1
            print(f"ISSN: {issn} | Periódico: {nome} | Extrato Qualis: {qualis}")

    # Gráfico
    plt.figure(figsize=(10, 6))
    barras = plt.bar(resumo.keys(), resumo.values(), color="skyblue", edgecolor="black")
    plt.title("Quantidade de Periódicos Encontrados por Estrato Qualis")
    plt.xlabel("Estrato Qualis")
    plt.ylabel("Quantidade")
    plt.grid(axis="y", linestyle="--", alpha=0.7)

    # Adicionar legenda com valores em cada barra
    for barra in barras:
        altura = barra.get_height()
        plt.text(barra.get_x() + barra.get_width() / 2, altura + 0.1, str(int(altura)),
                 ha="center", va="bottom")

    grafico_path = "grafico_qualis.png"
    plt.savefig(grafico_path)
    plt.show()

    # Gerar PDF
    gerar_pdf(relatorio, resumo, grafico_path)
    print("\n✅ Relatório gerado com sucesso: Relatorio_ISSN.pdf")


if __name__ == "__main__":
    main()
