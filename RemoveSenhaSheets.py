import openpyxl
from tkinter import Tk, filedialog


def remover_protecao_planilhas_xlsx():
    # Oculta a janela principal do tkinter
    root = Tk()
    root.withdraw()

    # Abre janela para selecionar o arquivo .xlsx
    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo Excel (.xlsx)",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not caminho:
        print("Nenhum arquivo selecionado.")
        return

    print(f"[✔] Abrindo: {caminho}")
    wb = openpyxl.load_workbook(caminho)

    for sheet in wb.worksheets:
        if sheet.protection.sheet:
            sheet.protection.sheet = False
            print(f"[✔] Proteção removida da planilha: {sheet.title}")

        # Desbloqueia todas as células da planilha
        for row in sheet.iter_rows():
            for cell in row:
                cell.protection = openpyxl.styles.Protection(locked=False)

    # Salva com novo nome para segurança
    novo_arquivo = caminho.replace(".xlsx", "_desprotegido.xlsx")
    wb.save(novo_arquivo)

    print(f"[✔] Arquivo salvo sem proteção: {novo_arquivo}")


if __name__ == "__main__":
    remover_protecao_planilhas_xlsx()
