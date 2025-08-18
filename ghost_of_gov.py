import asyncio
from playwright.async_api import async_playwright
import subprocess
import sys
import os

# --- BANNER / CORES (opcional) ---
try:
    import pyfiglet  # para ASCII art
except Exception:
    pyfiglet = None

try:
    import colorama  # habilita ANSI no Windows
    colorama.just_fix_windows_console()
except Exception:
    pass

RED = '\033[91m'
RESET = '\033[0m'

def print_banner():
    """Imprime o banner ASCII em vermelho, com fallback simples."""
    try:
        if pyfiglet is not None:
            banner_text = pyfiglet.figlet_format("Ghost of Gov", font="standard", width=200)
        else:
            banner_text = "GOV AUTOMATION"
    except Exception:
        banner_text = "GOV AUTOMATION"
    print("\n" + RED + banner_text + RESET)

# --- CONFIGURAÇÃO ---
CHROME_PATH_WINDOWS = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
REMOTE_DEBUGGING_PORT = 9222
INITIAL_WAIT_SECONDS = 35
SECONDARY_WAIT_SECONDS = 20
TRANSITION_WAIT_SECONDS = 10
PDF_LOAD_WAIT_SECONDS = 10
START_YEAR = 2000
END_YEAR = 2025
DOWNLOAD_DIRECTORY = os.path.join(os.getcwd(), "holerites_baixados")
TYPING_DELAY_MS = 50
POST_SEARCH_DELAY_MS = 500
# --- FIM DA CONFIGURAÇÃO ---

def launch_chrome_for_manual_login():
    """Inicia uma instância do Chrome com a porta de depuração aberta para o login manual."""
    if not os.path.exists(DOWNLOAD_DIRECTORY):
        os.makedirs(DOWNLOAD_DIRECTORY)
        print(f"Pasta de download criada em: {DOWNLOAD_DIRECTORY}")
    else:
        print(f"Os downloads serão salvos em: {DOWNLOAD_DIRECTORY}")

    profile_path = os.path.join(os.getcwd(), "hybrid_chrome_profile")
    command = [
        CHROME_PATH_WINDOWS,
        f"--remote-debugging-port={REMOTE_DEBUGGING_PORT}",
        f"--user-data-dir={profile_path}",
        "https://www.sou.sp.gov.br/sou.sp"
    ]

    # >>> Banner aqui, antes das instruções <<<
    print_banner()

    print("=" * 70)
    print("A EXECUÇÃO FINAL - INSTRUÇÕES:")
    print("1. Uma nova janela do Chrome será aberta.")
    print("2. FAÇA O LOGIN MANUALMENTE até a página principal.")
    print("3. Volte para este terminal e pressione ENTER.")
    print("=" * 70)
    subprocess.Popen(command)

async def select_year(page):
    """BLOCO 1 (INTOCÁVEL): Lógica para selecionar o ano."""
    print("\n--- BLOCO 1: SELEÇÃO DE ANO ---")
    year_dropdown_id = "#s2id_sp_formfield_reference_year"
    await page.locator(f"{year_dropdown_id} a.select2-choice").click()
    search_input = page.locator("#select2-drop input.select2-input")
    await search_input.wait_for(state='visible', timeout=10000)
    for year_to_check in range(START_YEAR, END_YEAR + 1):
        year_str = str(year_to_check)
        await search_input.fill("")  # limpa antes
        await search_input.type(year_str, delay=TYPING_DELAY_MS)
        await page.wait_for_timeout(200)
        target_option = page.locator(f'div.select2-result-label:has-text("{year_str}")')
        try:
            await target_option.wait_for(state="visible", timeout=1000)
            print(f"Ano '{year_str}' encontrado. Clicando...")
            await target_option.click()
            print("--- SUCESSO DO BLOCO 1 ---")
            return year_str
        except Exception:
            # tenta próximo ano
            pass
    raise Exception(f"Nenhum ano entre {START_YEAR}-{END_YEAR} foi encontrado.")

async def select_months_and_download(page, context, selected_year):
    """BLOCO 2 (CORRIGIDO): Manipula a nova aba e clica no botão de download real."""
    print(f"\n--- BLOCO 2: SELEÇÃO E DOWNLOAD PARA {selected_year} ---")
    await page.wait_for_timeout(TRANSITION_WAIT_SECONDS * 1000)

    month_dropdown_id = "#s2id_sp_formfield_reference_date"
    
    for month_num in range(1, 13):
        print("-" * 20)
        await page.locator(month_dropdown_id).click()
        
        month_search_input = page.locator("#select2-drop input.select2-input")
        try:
            await month_search_input.wait_for(state='visible', timeout=5000)
        except Exception:
            print(f"ERRO: Caixa de busca de mês não apareceu para {month_num:02d}/{selected_year}. Pulando.")
            await page.locator("body").press("Escape")
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue

        month_str = f"{month_num:02d}/{selected_year}"
        print(f"Buscando pelo mês '{month_str}'...")
        await month_search_input.fill("")  # limpa antes
        await month_search_input.type(month_str, delay=TYPING_DELAY_MS)
        
        month_option = page.locator(f'div.select2-result-label:has-text("{month_str}")')
        
        try:
            await month_option.wait_for(state="visible", timeout=1500)
            print(f"Mês '{month_str}' encontrado. Selecionando...")
            await month_option.click()
        except Exception:
            print(f"Mês '{month_str}' não encontrado. Pulando.")
            if await page.locator("#select2-drop").is_visible():
                await page.locator("body").press("Escape")
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue
            
        print(f"Aguardando {PDF_LOAD_WAIT_SECONDS} segundos...")
        await page.wait_for_timeout(PDF_LOAD_WAIT_SECONDS * 2000)
        
        pdf_button = page.locator('button:has-text("Visualizar PDF")')
        try:
            await pdf_button.wait_for(state="visible", timeout=5000)
        except Exception:
            print(f"AVISO: Botão 'Visualizar PDF' não visível para {month_str}.")
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue

        if await pdf_button.is_enabled():
            print("Botão 'Visualizar PDF' habilitado. Abrindo nova aba...")
            
            # --- INÍCIO DA LÓGICA CORRETA ---
            async with context.expect_page() as new_page_info:
                await pdf_button.click()
            
            pdf_page = await new_page_info.value
            print(f"Nova aba aberta. Aguardando carregamento...")
            await pdf_page.wait_for_load_state('networkidle')
            
            # Prepara para o download que será iniciado a partir da nova aba
            async with pdf_page.expect_download() as download_info:
                # Clica no botão de download DENTRO do visualizador (shadow DOM)
                clicked = False
                for sel in ("css=pdf-viewer >>> #download", "css=viewer-toolbar >>> #download", "#download"):
                    try:
                        await pdf_page.locator(sel).click(timeout=4000)
                        print(f"Botão de download clicado via seletor: {sel}")
                        clicked = True
                        break
                    except Exception:
                        pass
                if not clicked:
                    raise Exception("Não foi possível acionar o botão de download no visualizador.")
            
            download = await download_info.value
            save_path = os.path.join(DOWNLOAD_DIRECTORY, f"holerite_{selected_year}_{month_num:02d}.pdf")
            await download.save_as(save_path)
            print(f"### SUCESSO: Download concluído. PDF salvo em {save_path} ###")
            
            await pdf_page.close()
            print("Aba do PDF fechada.")
            # --- FIM DA LÓGICA CORRETA ---

        else:
            print(f"AVISO: Botão 'Visualizar PDF' para o mês {month_str} está desabilitado.")
        
        await page.wait_for_timeout(POST_SEARCH_DELAY_MS)

    print("--- SUCESSO DO BLOCO 2 ---")

# ========= NOVO: selecionar ano específico sem mexer no BLOCO 1 =========
async def select_specific_year(page, year_str: str) -> bool:
    """Abre o Select2 de ano e tenta selecionar exatamente `year_str` (ex.: '2019')."""
    year_dropdown_id = "#s2id_sp_formfield_reference_year"
    dropdown = page.locator(f"{year_dropdown_id} a.select2-choice")
    await dropdown.click()
    drop = page.locator("#select2-drop")
    try:
        await drop.wait_for(state='visible', timeout=10000)
    except Exception:
        # tenta mais um clique para garantir que o dropdown abriu
        await dropdown.click()
        await drop.wait_for(state='visible', timeout=10000)

    search_input = drop.locator("input.select2-input")
    await search_input.wait_for(state='visible', timeout=10000)
    await search_input.fill("")
    await search_input.type(year_str, delay=TYPING_DELAY_MS)
    await page.wait_for_timeout(200)

    option = drop.locator(f'div.select2-result-label:has-text("{year_str}")')
    try:
        await option.wait_for(state="visible", timeout=1200)
        print(f"Ano '{year_str}' encontrado. Clicando...")
        await option.click()
        return True
    except Exception:
        print(f"Ano '{year_str}' não encontrado. Pulando.")
        try:
            if await drop.is_visible():
                await page.locator("body").press("Escape")
        except Exception:
            pass
        return False
# =======================================================================

async def run_final_execution_flow(page, context):
    """Executa a rotina final com a lógica de download correta."""
    # --- ROTINA DE ACESSO ---
    print("\nExecutando rotina de acesso...")
    await page.locator('a[aria-label="Vínculos"]').click()
    await page.locator('a.select2-choice:visible').first.click()
    await page.locator('div.select2-result-label:has-text("Aposentado")').click()
    await page.locator('div.modal-content button:has-text("Entrar")').click()
    await page.wait_for_selector('h3[title="Contracheque"]', state='visible', timeout=45000)
    await page.locator('h3[title="Contracheque"]').click()
    print("Acesso à área concluído.")
    
    # --- ESPERAS ---
    print(f"Esperando {INITIAL_WAIT_SECONDS}s...")
    await page.wait_for_timeout(INITIAL_WAIT_SECONDS * 1000)
    await page.locator('#s2id_sp_formfield_type_of_request a.select2-choice').click()
    await page.locator('div.select2-result-label:has-text("Anteriores")').click()
    print("Tipo de solicitação alterado.")
    print(f"Esperando {SECONDARY_WAIT_SECONDS}s...")
    await page.wait_for_timeout(SECONDARY_WAIT_SECONDS * 1000)
    print("Página pronta.")

    # --- EXECUÇÃO DOS BLOCOS ---
    # 1) Seleciona o primeiro ano disponível (mantém o BLOCO 1 intocado)
    selected_year = await select_year(page)
    await select_months_and_download(page, context, selected_year)

    # 2) Segue para os anos seguintes até END_YEAR, um a um
    try:
        start_next = int(selected_year) + 1
    except Exception:
        start_next = END_YEAR + 1  # se não conseguir converter, encerra o loop

    for year in range(start_next, END_YEAR + 1):
        print("\n" + "=" * 28)
        print(f"=== INICIANDO ANO {year} ===")
        print("=" * 28)
        ok = await select_specific_year(page, str(year))
        if ok:
            await select_months_and_download(page, context, str(year))
        else:
            print(f"Nenhuma referência para {year}. Seguindo para o próximo.")

    print("\n" + "=" * 30)
    print("!!! A EXECUÇÃO FINAL FOI CONCLUÍDA !!!")
    print("=" * 30)

async def main():
    launch_chrome_for_manual_login()
    input()
    
    print("\nTentando conectar ao navegador...")
    async with async_playwright() as p:
        page = None
        context = None
        try:
            browser = await p.chromium.connect_over_cdp(f"http://localhost:{REMOTE_DEBUGGING_PORT}")
            context = browser.contexts[0]
            page = context.pages[0]
            print("Conexão estabelecida!")
            await run_final_execution_flow(page, context)
        except Exception as e:
            print(f"\n--- FALHA NA EXECUÇÃO FINAL ---")
            print(f"ERRO: {e}")
            if page:
                screenshot_path = f"falha_execucao_final.png"
                await page.screenshot(path=screenshot_path)
                print(f"Diagnóstico salvo em: {screenshot_path}")

if __name__ == "__main__":
    if not os.path.exists(CHROME_PATH_WINDOWS):
        print(f"ERRO: Caminho do Chrome não encontrado em '{CHROME_PATH_WINDOWS}'.")
        sys.exit(1)
        
    asyncio.run(main())
    print("\nScript finalizado. Pressione ENTER para sair.")
    input()
