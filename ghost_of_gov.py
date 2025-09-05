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
    page.on("dialog", lambda d: asyncio.create_task(d.dismiss()))
    await disable_beforeunload(page)
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
        # garante que nenhum modal antigo esteja aberto
        await close_leave_modal_if_present(page)

        year_str = str(year_to_check)
        await search_input.fill("")  # limpa antes
        await search_input.type(year_str, delay=TYPING_DELAY_MS)
        await page.wait_for_timeout(200)

        target_option = page.locator(f'div.select2-result-label:has-text("{year_str}")')
        try:
            await target_option.wait_for(state="visible", timeout=1000)
            print(f"Ano '{year_str}' encontrado. Clicando...")
            await target_option.click()

            # desfoca e fecha o modal "Sair da página?" caso apareça
            await defocus_form(page)
            await close_leave_modal_if_present(page)

            print("--- SUCESSO DO BLOCO 1 ---")
            return year_str
        except Exception:
            pass

    raise Exception(f"Nenhum ano entre {START_YEAR}-{END_YEAR} foi encontrado.")

async def release_focus_to_pdf_button(page):
    """Fecha Select2/máscara, remove foco de Ano/Mês/Demonstrativo e dá foco no botão Visualizar PDF."""
    # 1) fecha overlays do Select2 (drop e mask)
    for _ in range(30):  # ~3s
        try:
            drop = page.locator('#select2-drop')
            mask = page.locator('#select2-drop-mask')
            open_drop = (await drop.count()) > 0 and await drop.is_visible()
            open_mask = (await mask.count()) > 0 and await mask.is_visible()
            if not open_drop and not open_mask:
                break
            await page.keyboard.press('Escape')
        except Exception:
            pass
        await page.wait_for_timeout(60)

    # 2) remove classes de "ativo" dos containers e blur no activeElement
    try:
        await page.evaluate("""
        () => {
          const ids = [
            '#s2id_sp_formfield_reference_year',
            '#s2id_sp_formfield_reference_date',
            '#s2id_sp_formfield_demonstrative'
          ];
          for (const sel of ids) {
            const cont = document.querySelector(sel);
            if (!cont) continue;
            cont.classList?.remove('select2-container-active','select2-dropdown-open','select2-container-open');
            const input = cont.querySelector('input.select2-input');
            if (input && input.blur) input.blur();
            const choice = cont.querySelector('a.select2-choice');
            if (choice) {
              choice.setAttribute('aria-expanded','false');
              if (choice.blur) choice.blur();
            }
          }
          if (document.activeElement && document.activeElement.blur) {
            document.activeElement.blur();
          }
        }
        """)
    except Exception:
        pass

    # 3) clique neutro pra garantir perda de foco (título/área fora do form)
    clicked_neutral = False
    for sel in ['h3[title="Contracheque"]', 'header', 'main', 'div.container', 'body']:
        try:
            loc = page.locator(sel)
            if await loc.count():
                await loc.first.click(timeout=300)
                clicked_neutral = True
                break
        except Exception:
            pass
    if not clicked_neutral:
        try:
            await page.mouse.click(5, 5)
        except Exception:
            pass

    # 4) dá foco ao botão (opcional, ajuda o navegador a “entender” o próximo clique)
    try:
        btn = page.locator('button:has-text("Visualizar PDF")')
        if await btn.count():
            await btn.scroll_into_view_if_needed()
            await btn.focus()
    except Exception:
        pass

    await page.wait_for_timeout(60)
    
# Fecha o modal "Sair da página?" se aparecer
async def close_leave_modal_if_present(page):
    try:
        title = page.locator('h1#modal-title:has-text("Sair da página?")').first
        if await title.count() and await title.is_visible():
            dialog = title.locator('xpath=ancestor::div[contains(@class,"modal-dialog")]').first
            cancel_btn = dialog.locator('button:has-text("Cancelar")').first
            if await cancel_btn.count() and await cancel_btn.is_visible():
                await cancel_btn.click()
                try:
                    await dialog.wait_for(state='hidden', timeout=2000)
                except Exception:
                    pass
                return True
            # fallback: tecla ESC
            try:
                await page.keyboard.press("Escape")
            except Exception:
                pass
            return True
    except Exception:
        pass
    return False

# Desfoca sem clicar em links/botões (evita disparar o modal)
DEFOCUS_SETTLE_MS = 80  # 50–200 ms, ajuste se precisar

async def defocus_form(page):
    # fecha qualquer Select2/máscara que tenha ficado aberto
    try:
        await page.keyboard.press("Escape")
        await page.keyboard.press("Escape")
    except Exception:
        pass

    # tira o foco do elemento ativo
    try:
        await page.evaluate("document.activeElement && document.activeElement.blur()")
    except Exception:
        pass

    # clique neutro (evita links/botões que disparam modal)
    clicked = False
    for sel in [
        'main', '#sp-main-content', '.portlet .portlet-content',
        'form', 'div.container', 'div.content', 'body'
    ]:
        try:
            neutral = page.locator(f'{sel} >> :not(a,button,input,select,textarea,label)')
            if await neutral.count():
                await neutral.first.click(position={"x": 5, "y": 5}, timeout=300)
                clicked = True
                break
        except Exception:
            pass
    if not clicked:
        try:
            vp = await page.evaluate('() => ({w: innerWidth, h: innerHeight})')
            await page.mouse.click(max(1, vp["w"] - 5), max(1, vp["h"] - 5))
        except Exception:
            pass

    await page.wait_for_timeout(DEFOCUS_SETTLE_MS)
    
async def disable_beforeunload(page):
    try:
        await page.add_init_script("""
            (function() {
              window.onbeforeunload = null;
              const _addEventListener = window.addEventListener;
              window.addEventListener = function(type, listener, options) {
                if (type === 'beforeunload') return; // bloqueia novos
                return _addEventListener.call(this, type, listener, options);
              };
            })();
        """)
        await page.evaluate("""
            window.onbeforeunload = null;
            window.addEventListener = (function(orig){
              return function(type, listener, options){
                if(type === 'beforeunload') return;
                return orig.call(this, type, listener, options);
              };
            })(window.addEventListener);
        """)
    except Exception:
        pass    

async def select_months_and_download(page, context, selected_year):
    """Seleciona mês a mês e baixa os PDFs (normal e suplementar quando existir)."""
    print(f"\n--- BLOCO 2: SELEÇÃO E DOWNLOAD PARA {selected_year} ---")
    await page.wait_for_timeout(TRANSITION_WAIT_SECONDS * 1000)

    month_dropdown_id = "#s2id_sp_formfield_reference_date"

    for month_num in range(1, 13):
        print("-" * 20)

        # Antes de abrir o Select2 do mês, garanta que não há modal aberto
        await close_leave_modal_if_present(page)

        await page.locator(month_dropdown_id).click()
        await close_leave_modal_if_present(page)  # se o clique disparar o modal, fecha

        month_search_input = page.locator("#select2-drop input.select2-input")
        try:
            await month_search_input.wait_for(state='visible', timeout=5000)
        except Exception:
            print(f"ERRO: Caixa de busca de mês não apareceu para {month_num:02d}/{selected_year}. Pulando.")
            try:
                await page.locator("body").press("Escape")
            except Exception:
                pass
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue

        month_str = f"{month_num:02d}/{selected_year}"
        print(f"Buscando pelo mês '{month_str}'...")
        await month_search_input.fill("")
        await month_search_input.type(month_str, delay=TYPING_DELAY_MS)

        month_option = page.locator(f'div.select2-result-label:has-text("{month_str}")')
        try:
            await month_option.wait_for(state="visible", timeout=1500)
            print(f"Mês '{month_str}' encontrado. Selecionando...")
            await month_option.click()
            await close_leave_modal_if_present(page)  # se aparecer ao selecionar, fecha
        except Exception:
            print(f"Mês '{month_str}' não encontrado. Pulando.")
            if await page.locator("#select2-drop").is_visible():
                try:
                    await page.locator("body").press("Escape")
                except Exception:
                    pass
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue

        # Desfoca imediatamente o Select2 (evita borda azul/scroll)
        try:
            await defocus_form(page)
            await page.locator('#select2-drop').wait_for(state='hidden', timeout=2000)
        except Exception:
            pass
        await close_leave_modal_if_present(page)

        print(f"Aguardando {PDF_LOAD_WAIT_SECONDS} segundos...")
        await page.wait_for_timeout(PDF_LOAD_WAIT_SECONDS * 2000)

        # Garante foco longe dos selects e modal fechado
        await release_focus_to_pdf_button(page)
        await close_leave_modal_if_present(page)

        baixou_por_demonstrativo = False

        # 1) Abrir "Demonstrativo" (se houver) e baixar todas as opções
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
                    await close_leave_modal_if_present(page)
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

            if labels:
                for label in labels:
                    # reabre se necessário
                    if not await page.locator('#select2-drop').is_visible():
                        reopened = False
                        for sel in [
                            '#s2id_sp_formfield_demonstrative a.select2-choice',
                            'xpath=//label[contains(normalize-space(.),"Demonstrativo")]/following::a[contains(@class,"select2-choice")][1]',
                            'xpath=//span[contains(@aria-label,"Demonstrativo")]//a[contains(@class,"select2-choice")]',
                            'xpath=//span[contains(@aria-label,"Demonstrativo")]'
                        ]:
                            try:
                                loc = page.locator(sel)
                                if await loc.count():
                                    await loc.first.click(timeout=2000)
                                    await close_leave_modal_if_present(page)
                                    reopened = True
                                    break
                            except Exception:
                                pass
                        if not reopened:
                            print("AVISO: não consegui reabrir o dropdown de Demonstrativo.")
                            break
                        await page.locator('#select2-drop').wait_for(state='visible', timeout=4000)

                    option = page.locator('#select2-drop li.select2-result-selectable').filter(has_text=label).first
                    try:
                        await option.scroll_into_view_if_needed()
                        await option.click(timeout=4000)
                        await close_leave_modal_if_present(page)
                    except Exception:
                        print(f"AVISO: não consegui clicar em '{label}'. Pulando.")
                        continue

                    # Desfoca o demonstrativo e prepara o botão
                    try:
                        await defocus_form(page)
                        await page.locator('#select2-drop').wait_for(state='hidden', timeout=2000)
                    except Exception:
                        pass
                    await release_focus_to_pdf_button(page)
                    await close_leave_modal_if_present(page)

                    # aguarda habilitar
                    pdf_button = page.locator('button:has-text("Visualizar PDF")')
                    for _ in range(16):
                        try:
                            if await pdf_button.is_enabled():
                                break
                        except Exception:
                            pass
                        await page.wait_for_timeout(500)
                    if not await pdf_button.is_enabled():
                        print(f"AVISO: Botão 'Visualizar PDF' segue desabilitado após '{label}'.")
                        continue

                    await pdf_button.scroll_into_view_if_needed()
                    await page.wait_for_timeout(50)

                    print(f"Baixando demonstrativo: {label}")
                    async with context.expect_page() as new_page_info:
                        try:
                            await pdf_button.click()
                        except Exception:
                            await pdf_button.click(force=True)
                    pdf_page = await new_page_info.value

                    try:
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

                        sanitized = re.sub(r'[^A-Za-z0-9_-]+', '_', label).strip('_')
                        save_path = os.path.join(DOWNLOAD_DIRECTORY, f"holerite_{selected_year}_{month_num:02d}_{sanitized}.pdf")
                        await download.save_as(save_path)
                        print(f"### SUCESSO: Download concluído. PDF salvo em {save_path} ###")
                        baixou_por_demonstrativo = True
                    except Exception as e:
                        print(f"AVISO: Falha ao baixar demonstrativo '{label}': {e}")
                    finally:
                        try:
                            await pdf_page.close()
                            await page.bring_to_front()
                            print("Aba do PDF (demonstrativo) fechada.")
                        except Exception:
                            pass

                if baixou_por_demonstrativo:
                    await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
                    continue
            else:
                try:
                    await page.keyboard.press("Escape")
                except Exception:
                    pass

        # 2) Fluxo sem demonstrativo
        pdf_button = page.locator('button:has-text("Visualizar PDF")')
        try:
            await pdf_button.wait_for(state="visible", timeout=5000)
        except Exception:
            print(f"AVISO: Botão 'Visualizar PDF' não visível para {month_str}.")
            await page.wait_for_timeout(POST_SEARCH_DELAY_MS)
            continue

        if await pdf_button.is_enabled():
            # reforços contra modal + foco
            try:
                await defocus_form(page)
                await page.locator('#select2-drop').wait_for(state='hidden', timeout=2000)
            except Exception:
                pass
            await release_focus_to_pdf_button(page)
            await close_leave_modal_if_present(page)

            await pdf_button.scroll_into_view_if_needed()
            await page.wait_for_timeout(50)

            print("Botão 'Visualizar PDF' habilitado. Abrindo nova aba...")
            async with context.expect_page() as new_page_info:
                try:
                    await pdf_button.click()
                except Exception:
                    await pdf_button.click(force=True)
            pdf_page = await new_page_info.value

            print("Nova aba aberta. Aguardando carregamento...")
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
            await page.bring_to_front()
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

    # garante que não há modal aberto antes de interagir
    await close_leave_modal_if_present(page)

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

        # desfoca e fecha o modal "Sair da página?" caso tenha surgido
        await defocus_form(page)
        await close_leave_modal_if_present(page)

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

    # Seleciona "Anteriores" e já trata o modal "Sair da página?"
    await page.locator('#s2id_sp_formfield_type_of_request a.select2-choice').click()
    await page.locator('div.select2-result-label:has-text("Anteriores")').click()
    await page.wait_for_timeout(50)  # respiro pro Select2 fechar
    await close_leave_modal_if_present(page)
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

            # --- timeouts padrão (recomendado) ---
            context.set_default_timeout(45000)
            page.set_default_timeout(45000)

            # --- limpeza preventiva de qualquer modal pendente ---
            await close_leave_modal_if_present(page)

            # --- Handler global para qualquer dialog (alert/confirm/prompt/beforeunload) ---
            async def _dismiss_dialog(dlg):
                try:
                    await dlg.dismiss()
                except Exception:
                    pass

            page.on("dialog", _dismiss_dialog)

            # --- Qualquer nova aba/página (ex.: visualizador de PDF) também herda o handler ---
            def _wire_new_page(pg):
                pg.on("dialog", lambda d: asyncio.create_task(_dismiss_dialog(d)))
            context.on("page", _wire_new_page)

            print("Conexão estabelecida!")
            await run_final_execution_flow(page, context, vinculo_choice, years_mode)

        except Exception as e:
            print(f"\n--- FALHA NA EXECUÇÃO FINAL ---")
            print(f"ERRO: {e}")
            if page:
                screenshot_path = "falha_execucao_final.png"
                try:
                    await page.screenshot(path=screenshot_path)
                    print(f"Diagnóstico salvo em: {screenshot_path}")
                except Exception:
                    pass

if __name__ == "__main__":
    if not os.path.exists(CHROME_PATH_WINDOWS):
        print(f"ERRO: Caminho do Chrome não encontrado em '{CHROME_PATH_WINDOWS}'.")
        sys.exit(1)

    asyncio.run(main())
    print("\nScript finalizado. Pressione ENTER para sair.")
    input()
