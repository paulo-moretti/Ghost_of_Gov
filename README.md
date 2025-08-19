# Ghost of Gov

## Descrição do Projeto

O `Ghost of Gov` é uma automação desenvolvido em Python para automatizar o processo de download de holerites (contracheques) do portal `SouGov.sp.gov.br`. Uma atualização das automações Android utilizadas anteriromente no Galaxye s10+ e Redmi 14C... Atualmente rodando na web, utilizando a biblioteca `Playwright`, ele simula a interação de um usuário com o navegador para navegar pelas páginas, selecionar anos e meses, e baixar os arquivos PDF dos holerites. O script é projetado para ser executado após um login manual inicial no portal, garantindo a segurança das credenciais do usuário e a conformidade com os termos de uso do site.

Este projeto visa simplificar a obtenção de documentos financeiros para servidores públicos do estado de São Paulo, eliminando a necessidade de downloads manuais repetitivos.

## Funcionalidades

- **Automação de Download**: Baixa holerites em PDF de forma automatizada, com lógica aprimorada para aguardar o carregamento do PDF e acionar o botão de download.
- **Seleção de Vínculo**: Permite ao usuário escolher entre diferentes vínculos (ex: SPPREV Aposentado, SE Inativo) diretamente pelo terminal.
- **Seleção Flexível de Anos**: Oferece opções para baixar holerites de todos os anos disponíveis (dentro de um intervalo configurável) ou de um intervalo de anos específico definido pelo usuário.
- **Navegação Inteligente**: Interage com dropdowns de ano e mês para selecionar os períodos desejados.
- **Gerenciamento de Sessão**: Utiliza um perfil de navegador persistente (`hybrid_chrome_profile`) para manter o estado de login.
- **Tratamento de Erros Aprimorado**: Inclui lógica para lidar com meses ou anos não encontrados, botões desabilitados e captura de screenshot em caso de falha na execução final para diagnóstico.
- **Configurável**: Permite ajustar o diretório de download, anos de início e fim, e tempos de espera.
- **Interface Amigável**: Exibe um banner ASCII (se `pyfiglet` estiver instalado) e instruções claras para a interação manual inicial.

## Requisitos

Para executar este script, você precisará:

- **Python 3.x**: Certifique-se de ter o Python instalado em seu sistema.
- **Google Chrome**: O script utiliza o Google Chrome para automação. O caminho padrão é `C:\Program Files\Google\Chrome\Application\chrome.exe` no Windows, mas pode ser alterado na configuração.
- **Bibliotecas Python**: As seguintes bibliotecas Python são necessárias:
    - `playwright`: Para automação do navegador.
    - `pyfiglet` (opcional): Para exibir o banner ASCII. Se não estiver instalado, um banner simples será usado.
    - `colorama` (opcional): Para cores no console, especialmente útil no Windows.

Você pode instalar as bibliotecas Python usando `pip`:

```bash
pip install playwright pyfiglet colorama
playwright install
```

**Nota**: O comando `playwright install` é crucial para baixar os drivers do navegador necessários.

## Como Usar

1.  **Clone o Repositório (ou baixe o arquivo `ghost_of_gov.py`):**

    ```bash
    git clone https://github.com/paulo-moretti/ghost_of_gov.git 
    cd ghost_of_gov
    ```

2.  **Instale as Dependências:**

    ```bash
pip install playwright pyfiglet colorama
playwright install
    ```

3.  **Execute o Script:**

    ```bash
    python ghost_of_gov.py
    ```

4.  **Login Manual:**

    -   Uma nova janela do Chrome será aberta.
    -   **FAÇA O LOGIN MANUALMENTE** no portal `https://www.sou.sp.gov.br/sou.sp` até a página principal onde os holerites são exibidos.
    -   Após o login, volte para o terminal onde o script está rodando e pressione `ENTER`.

5.  **Seleção de Vínculo e Anos:**

    -   O script perguntará qual vínculo você deseja usar (SPPREV Aposentado ou SE Inativo).
    -   Em seguida, você poderá escolher entre baixar holerites de todos os anos disponíveis ou especificar um intervalo de anos.

6.  **Automação:**

    -   O script assumirá o controle do navegador e começará a baixar os holerites para os anos e meses selecionados.
    -   Os arquivos PDF serão salvos na pasta `holerites_baixados` dentro do diretório do projeto.

## Configuração

Você pode ajustar as seguintes variáveis no arquivo `ghost_of_gov.py` para personalizar o comportamento do script:

-   `CHROME_PATH_WINDOWS`: Caminho para o executável do Chrome. Altere se o seu Chrome estiver em um local diferente ou se estiver usando Linux/macOS (neste caso, o Playwright geralmente encontra automaticamente, mas pode ser especificado).
-   `REMOTE_DEBUGGING_PORT`: Porta usada para depuração remota do Chrome. Padrão é `9222`.
-   `INITIAL_WAIT_SECONDS`: Tempo de espera inicial após o login manual (em segundos).
-   `SECONDARY_WAIT_SECONDS`: Tempo de espera secundário após algumas interações (em segundos).
-   `TRANSITION_WAIT_SECONDS`: Tempo de espera durante transições de página (em segundos).
-   `PDF_LOAD_WAIT_SECONDS`: Tempo de espera para o carregamento do PDF antes de tentar o download (em segundos).
-   `START_YEAR`: Ano inicial para buscar holerites.
-   `END_YEAR`: Ano final para buscar holerites.
-   `DOWNLOAD_DIRECTORY`: Diretório onde os holerites serão salvos. Padrão é `holerites_baixados` na pasta do script.
-   `TYPING_DELAY_MS`: Atraso em milissegundos entre cada caractere digitado (simula digitação humana).
-   `POST_SEARCH_DELAY_MS`: Atraso em milissegundos após uma busca ou interação.

## Observações Importantes

-   Este script foi desenvolvido para o portal `sou.sp.gov.br`. Alterações na interface do site podem exigir atualizações no código.
-   O login inicial é manual por questões de segurança e para evitar o armazenamento de credenciais no script.
-   O script cria um perfil de usuário do Chrome (`hybrid_chrome_profile`) para persistir o estado de login e outras configurações do navegador. Não exclua esta pasta se quiser manter o login e as configurações.
-   A automação de sites pode ser sensível a mudanças. Se o script parar de funcionar, verifique se houve atualizações no site do SouGov.sp.gov.br.
-   Em caso de falha na execução, uma screenshot será salva na pasta do projeto para auxiliar no diagnóstico.

---

© 2025 Paulo Eduardo Moretti. Todos os direitos reservados.


