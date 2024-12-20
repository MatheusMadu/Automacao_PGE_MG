# Automação de Consultas - PGE MG

Este projeto é uma ferramenta para automatizar consultas no site da Procuradoria Geral do Estado de Minas Gerais (PGE-MG). A aplicação utiliza Selenium para interagir com o site, manipulação de planilhas Excel com OpenPyXL e uma interface gráfica intuitiva criada com Tkinter.

## Funcionalidades
- Automação de consultas de débitos no site da PGE-MG.
- Preenchimento de planilhas Excel com os resultados das consultas.
- Geração de PDFs dinâmicos para cada consulta realizada.
- Interface gráfica com:
  - Seleção de arquivo Excel e diretório de saída.
  - Nomeação personalizada para demandas.
  - Barra de progresso e área de logs para acompanhamento em tempo real.
  - Informações de status, incluindo tempo estimado de conclusão (ETA).
- Geração de relatório final com resumo do processo.

## Pré-requisitos
- Python 3.8 ou superior.
- Dependências listadas no arquivo `requirements.txt`:
  - `selenium`
  - `webdriver-manager`
  - `openpyxl`
  - `tkinter` (padrão em instalações do Python para Windows).
- Navegador Google Chrome e Chromedriver instalados.

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu_usuario/automacao_pge_mg.git
   cd automacao_pge_mg
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Certifique-se de que o Chromedriver está configurado no `PATH` ou na mesma pasta do projeto.

## Uso

1. Execute o programa:
   ```bash
   python nome_do_script.py
   ```

2. Na interface gráfica:
   - Escolha o arquivo Excel contendo os dados de consulta.
   - Insira um nome para a demanda.
   - Escolha o diretório onde os PDFs e resultados serão salvos.
   - Clique em **Iniciar Processo** para começar.

## Estrutura do Projeto

- **`iniciar_processo`**: Função principal que executa as consultas e preenche os resultados no Excel.
- **`gerar_pdf_dinamico`**: Gera PDFs das páginas de resultados das consultas.
- **`main`**: Cria a interface gráfica do usuário.
- **`configurar_chrome_options`**: Configura o navegador para automação.

## Principais Tecnologias
- **[Selenium](https://www.selenium.dev/)**: Automação de interação com navegadores.
- **[Tkinter](https://docs.python.org/3/library/tkinter.html)**: Interface gráfica para aplicativos Python.
- **[OpenPyXL](https://openpyxl.readthedocs.io/)**: Manipulação de arquivos Excel.
- **[Webdriver Manager](https://github.com/SergeyPirogov/webdriver_manager)**: Gerenciamento automático do Chromedriver.

## Contato
- **Autor**: Matheus Madureira da Fonseca
- **E-mail**: madureira-matheus@hotmail.com
