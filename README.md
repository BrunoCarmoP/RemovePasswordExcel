# 🔓 Sheet Unlocker - Ferramenta de Desproteção de Planilhas Excel em Python

[![Feito com Python](https://img.shields.io/badge/Feito%20com-Python-blue)](https://www.python.org/)
[![Licença MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](LICENSE)

Um script **focado, leve e eficiente** desenvolvido em Python para **remover a proteção de planilhas e o bloqueio de células** em arquivos do Microsoft Excel (`.xlsx`). O projeto é ideal para automatizar o desbloqueio de arquivos próprios onde a senha foi esquecida ou em fluxos de trabalho que exigem a edição rápida de planilhas protegidas.

---

## ✨ Funcionalidades e Mecanismo

O script concentra-se em uma única e crucial tarefa: desabilitar as restrições de edição de planilhas Excel.

### 1. Seleção de Arquivo e Interface (GUI)
* **Módulo:** `tkinter`
* A função principal inicia uma interface de diálogo (`filedialog.askopenfilename`) para permitir que o usuário selecione visualmente o arquivo `.xlsx`, tornando a ferramenta acessível a usuários sem familiaridade com a linha de comando.
* A janela principal do `tkinter` é suprimida (`root.withdraw()`) para manter a interface limpa.

### 2. Processamento Principal (openpyxl)
* **Módulo:** `openpyxl`
* Após a seleção, o arquivo é carregado na memória.
* **Desproteção de Planilhas:** O script itera sobre cada folha de trabalho (`wb.worksheets`). Se o atributo **`sheet.protection.sheet`** for detectado como `True` (protegido), ele é definido como `False`. Isso remove a senha e a proteção geral da planilha.
* **Desbloqueio de Células:** O Excel usa um atributo de "bloqueio" em células individuais para determinar quais podem ser editadas após a desproteção da folha. O script garante a edição total ao iterar sobre **todas as células** e definir a propriedade **`cell.protection.locked`** como `False`.

### 3. Saída Segura
* O arquivo modificado é sempre salvo com o sufixo **`_desprotegido.xlsx`** (ex: `dados.xlsx` $\rightarrow$ `dados_desprotegido.xlsx`).
* **O arquivo original é mantido intacto**, prevenindo perda de dados em caso de falha.

---

## 🛠️ Requisitos e Instalação

### Pré-requisitos
* **Python 3.x**

### Instalação das Dependências

O projeto requer a biblioteca `openpyxl` para manipulação de arquivos Excel.

```bash

pip install openpyxl
```


### Desenvolvido por @BrunoCarmoP


