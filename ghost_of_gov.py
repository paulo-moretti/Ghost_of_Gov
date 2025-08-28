import asyncio
from playwright.async_api import async_playwright
import subprocess
import sys
import os
import re

# BANNER / CORES
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

# CONFIGURAÇÃO 
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
# FIM DA CONFIGURAÇÃO

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

    # Banner
    print_banner()

    print("=" * 70)
    print("A EXECUÇÃO FINAL - INSTRUÇÕES:")
    print("1. Uma nova janela do Chrome será aberta.")
    print("2. FAÇA O LOGIN MANUALMENTE até a página principal.")
    print("3. Volte para este terminal e pressione ENTER.")
    print("=" * 70)
    subprocess.Popen(command)

# Escolha do vínculo no terminal  #
def ask_vinculo_from_terminal() -> str:
    """Pergunta ao usuário qual vínculo usar: 1-SPPREV (Aposentado) | 2-SE (Inativo)."""
    print("\nEscolha o vínculo para usar:")
    print("1 - SPPREV (Aposentado)")
    print("2 - SE (Inativo)")
    while True:
        choice = input("Digite 1 ou 2 e pressione ENTER: ").strip()
        if choice in ("1", "2"):
            return choice
        print("Opção inválida. Tente novamente (1 ou 2).")

async def choose_vinculo(page, choice: str):
    """
    Abre o seletor de vínculo (Select2) e clica no item correspondente SEM digitar.
    choice = '1' -> SPPREV (Aposentado)
    choice = '2' -> SE (Inativo)
    """
    await page.locator('a[aria-label="Vínculos"]').click()
    await page.locator('a.select2-choice:visible').first.click()

    drop = page.locator('#select2-drop')
    await drop.wait_for(state='visible', timeout=10000)
    await drop.locator('div.select2-result-label').first.wait_for(state='visible', timeout=10000)

    clicked = False
    if choice == "1":
        try:
            await drop.locator('div.select2-result-label:has-text("Aposentado")').first.click(timeout=2000)
            clicked = True
        except Exception:
            pass
    else:
        try:
            await drop.locator('div.select2-result-label').filter(has_text="SE").filter(has_text="Inativo").first.click(timeout=2500)
            clicked = True
        except Exception:
            # Fallback: varre rótulos procurando SE + Inativo
            labels = drop.locator('div.select2-result-label')
            count = await labels.count()
            for i in range(count):
                try:
                    text = await labels.nth(i).inner_text()
                    if "SE" in text and "Inativo" in text:
                        await labels.nth(i).click()
                        clicked = True
                        break
                except Exception:
                    pass

    if not clicked:
        
        try:
            await drop.locator('li.select2-result-selectable.select2-highlighted div.select2-result-label').click(timeout=1500)
            clicked = True
        except Exception:
            pass

    if not clicked:
        raise Exception("Não foi possível selecionar o vínculo desejado no Select2.")

    await page.locator('div.modal-content button:has-text("Entrar")').click()
    await page.wait_for_selector('h3[title="Contracheque"]', state='visible', timeout=45000)
    await page.locator('h3[title="Contracheque"]').click()
    print("Acesso à área concluído.")

# seleção de anos via terminal #
def ask_years_mode_from_terminal():
    """
    Retorna:
      {'mode': 'all'}                        -> baixar todos (START_YEAR..END_YEAR)
      {'mode': 'range', 'start': AAAA, 'end': AAAA} -> baixar intervalo
    """
    print("\nSeleção de anos:")
    print("1 - Informar intervalo (de XXXX a XXXX)")
    print("2 - Baixar TODOS os anos disponíveis")
    while True:
        opt = input("Escolha 1 ou 2 e pressione ENTER: ").strip()
        if opt == "2":
            return {"mode": "all"}
        if opt == "1":
            while True:
                try:
                    y1 = int(input("De (AAAA): ").strip())
                    y2 = int(input("Até (AAAA): ").strip())
                    if y1 > y2:
                        y1, y2 = y2, y1
                    return {"mode": "range", "start": y1, "end": y2}
                except Exception:
                    print("Valores inválidos. Tente novamente (somente números AAAA).")
        print("Opção inválida. Tente novamente.")

async def select_year(page):
    """Lógica para selecionar o ano."""
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
            pass
    raise Exception(f"Nenhum ano entre {START_YEAR}-{END_YEAR} foi encontrado.")

async def select_months_and_download(page, context, selected_year):
    """Manipula a nova aba e clica no botão de download real."""
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

        baixou_por_demonstrativo = False

        # 1) Tenta abrir o dropdown "Demonstrativo" (Select2)
        demonstrativo_opened = False
        for sel in [
            '#s2id_sp_formfield_demonstrative a.select2-choice',
            'xpath=//label[contains(normalize-space(.),"Demonstrativo")]/following::a[contains(@class,"select2-choice")][1]',
            'xpath=//span[contains(@aria-label,"Demonstrativo")]//a[contains(@class,"select2-choice")]',
            'xpath=//span[contains(@aria-label,"Demonstrativo")]'
        ]:
            try:
                loc = page.locator(sel)
                if await loc.count() > 0:
                    await loc.first.click(timeout=2000)
                    demonstrativo_opened = True
                    break
            except Exception:
                pass

        if demonstrativo_opened:
            drop = page.locator('#select2-drop')
            try:
                await drop.wait_for(state='visible', timeout=4000)
            except Exception:
                pass

            # 2) Coleta as opções válidas
            labels = []
            try:
                items = drop.locator('li.select2-result-selectable')
                n = await items.count()
                for i in range(n):
                    try:
                        t = (await items.nth(i).inner_text()).strip()
                        if t and "-- Nenhum --" not in t:
                            labels.append(t)
                    except Exception:
                        pass
            except Exception:
                labels = []

            # 3) Se tiver, seleciona e baixa TODAS as opções (normal/suplementar)
            if labels:
                for idx, label in enumerate(labels, start=1):

                    # Garante que o dropdown esteja ABERTO: se já estiver visível, não tenta reabrir (evita toggle fechar)
                    if not await page.locator('#select2-drop').is_visible():
                        opened = False
                        for sel in [
                            '#s2id_sp_formfield_demonstrative a.select2-choice',
                            'xpath=//label[contains(normalize-space(.),"Demonstrativo")]/following::a[contains(@class,"select2-choice")][1]',
                            'xpath=//span[contains(@aria-label,"Demonstrativo")]//a[contains(@class,"select2-choice")]',
                            'xpath=//span[contains(@aria-label,"Demonstrativo")]'
                        ]:
                            try:
                                loc = page.locator(sel)
                                if await loc.count() > 0:
                                    await loc.first.click(timeout=2000)
                                    opened = True
                                    break
                            except Exception:
                                pass
                        if not opened:
                            print("AVISO: não consegui abrir o dropdown de Demonstrativo.")
                            break

                        await page.locator('#select2-drop').wait_for(state='visible', timeout=4000)

                    # Clica na opção pelo texto (usa o label que coletamos)  << CORRIGIDO AQUI
                    option = page.locator('#select2-drop li.select2-result-selectable').filter(has_text=label).first
                    try:
                        await option.scroll_into_view_if_needed()
                        await option.click(timeout=4000)
                    except Exception:
                        print(f"AVISO: não consegui clicar em '{label}'. Pulando.")
                        continue

                    # Aguarda o botão habilitar após a seleção (algumas telas demoram um pouco)
                    pdf_button = page.locator('button:has-text("Visualizar PDF")')
                    for _ in range(16):  # ~8s (16 * 500ms)
                        try:
                            if await pdf_button.is_enabled():
                                break
                        except Exception:
                            pass
                        await page.wait_for_timeout(500)

                    if not await pdf_button.is_enabled():
                        print(f"AVISO: Botão 'Visualizar PDF' segue desabilitado após '{label}'.")
                        continue

                    # Nome do arquivo com o rótulo sanitizado
                    sanitized = re.sub(r'[^A-Za-z0-9_-]+', '_', label).strip('_')
                    save_path = os.path.join(
                        DOWNLOAD_DIRECTORY,
                        f"holerite_{selected_year}_{month_num:02d}_{sanitized}.pdf"
                    )

                    print(f"Baixando demonstrativo: {label}")
                    async with context.expect_page() as new_page_info:
                        await pdf_button.click()
                    pdf_page = await new_page_info.value

                    try:
                        await pdf_page.wait_for_load_state('networkidle')
                        async with pdf_page.expect_download() as download_info:
                            clicked = False
                            for sel in (
                                "css=pdf-viewer >>> #download",
                                "css=viewer-toolbar >>> #download",
                                "#download"
                            ):
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
                        await download.save_as(save_path)
                        print(f"### SUCESSO: Download concluído. PDF salvo em {save_path} ###")
                        baixou_por_demonstrativo = True
                    except Exception as e:
                        print(f"AVISO: Falha ao baixar demonstrativo '{label}': {e}")
                    finally:
                        try:
                            await pdf_page.close()
                            print("Aba do PDF (demonstrativo) fechada.")
                        except Exception:
                            pass

                # se baixou algum demonstrativo, passa direto ao próximo mês
                if baixou_por_demonstrativo:
                    await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
                    continue
            else:
                # não havia opções úteis; fecha dropdown se ficou aberto
                try:
                    await page.keyboard.press("Escape")
                except Exception:
                    pass

        pdf_button = page.locator('button:has-text("Visualizar PDF")')
        try:
            await pdf_button.wait_for(state="visible", timeout=5000)
        except Exception:
            print(f"AVISO: Botão 'Visualizar PDF' não visível para {month_str}.")
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue

        if await pdf_button.is_enabled():
            print("Botão 'Visualizar PDF' habilitado. Abrindo nova aba...")
            async with context.expect_page() as new_page_info:
                await pdf_button.click()
            
            pdf_page = await new_page_info.value
            print(f"Nova aba aberta. Aguardando carregamento...")
            await pdf_page.wait_for_load_state('networkidle')
            
            async with pdf_page.expect_download() as download_info:
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
        else:
            print(f"AVISO: Botão 'Visualizar PDF' para o mês {month_str} está desabilitado.")
        
        await page.wait_for_timeout(POST_SEARCH_DELAY_MS)

    print("--- SUCESSO DO BLOCO 2 ---")

# selecionar ano específico
async def select_specific_year(page, year_str: str) -> bool:
    """Abre o Select2 de ano e tenta selecionar exatamente `year_str`."""
    year_dropdown_id = "#s2id_sp_formfield_reference_year"
    dropdown = page.locator(f"{year_dropdown_id} a.select2-choice")
    await dropdown.click()
    drop = page.locator("#select2-drop")
    try:
        await drop.wait_for(state='visible', timeout=10000)
    except Exception:
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

async def run_final_execution_flow(page, context, vinculo_choice: str, years_mode: dict):
    """Executa a rotina final com a lógica de download correta."""
    # --- ROTINA DE ACESSO ---
    print("\nExecutando rotina de acesso...")
    await choose_vinculo(page, vinculo_choice)
    print(f"Vínculo selecionado: {'SPPREV (Aposentado)' if vinculo_choice=='1' else 'SE (Inativo)'}")
    
    # --- ESPERAS ---
    print(f"Esperando {INITIAL_WAIT_SECONDS}s...")
    await page.wait_for_timeout(INITIAL_WAIT_SECONDS * 1000)
    await page.locator('#s2id_sp_formfield_type_of_request a.select2-choice').click()
    await page.locator('div.select2-result-label:has-text("Anteriores")').click()
    print("Tipo de solicitação alterado.")
    print(f"Esperando {SECONDARY_WAIT_SECONDS}s...")
    await page.wait_for_timeout(SECONDARY_WAIT_SECONDS * 1000)
    print("Página pronta.")

    # --- EXECUÇÃO DOS BLOCOS conforme escolha de anos ---
    if years_mode.get("mode") == "all":
        # Encontra o primeiro ano disponível (BLOCO 1) e segue até END_YEAR
        selected_year = await select_year(page)
        await select_months_and_download(page, context, selected_year)
        try:
            start_next = int(selected_year) + 1
        except Exception:
            start_next = END_YEAR + 1
        for year in range(start_next, END_YEAR + 1):
            print("\n" + "=" * 28)
            print(f"=== INICIANDO ANO {year} ===")
            print("=" * 28)
            ok = await select_specific_year(page, str(year))
            if ok:
                await select_months_and_download(page, context, str(year))
            else:
                print(f"Nenhuma referência para {year}. Seguindo para o próximo.")
    else:
        y1, y2 = years_mode["start"], years_mode["end"]
        print(f"Baixando intervalo de anos: {y1} a {y2}")
        for year in range(y1, y2 + 1):
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
    input()  # aguarda login manual

    vinculo_choice = ask_vinculo_from_terminal()
    years_mode = ask_years_mode_from_terminal()

    print("\nTentando conectar ao navegador...")
    async with async_playwright() as p:
        page = None
        context = None
        try:
            browser = await p.chromium.connect_over_cdp(f"http://localhost:{REMOTE_DEBUGGING_PORT}")
            context = browser.contexts[0]
            page = context.pages[0]
            print("Conexão estabelecida!")
            await run_final_execution_flow(page, context, vinculo_choice, years_mode)
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
