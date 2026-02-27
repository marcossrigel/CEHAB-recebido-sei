from oauth2client.service_account import ServiceAccountCredentials
import gspread

from seleniumbase import SB
import os
import time

CAMINHO_CREDENCIAL = "formulariosolicitacaopagamento-6292734a5ede.json"
PLANILHA_ID = "1lkM9yOjhu_D2nQjRFl-Wt6lNgWPvzl2wbQiaO633-KM"
GID_BMS_2026 = 1189147903

STATUS_FILTRAR = "AGUARDANDO SEI"
STATUS_DESTINO = "RECEBIDO NO SEI"

SEI_LOGIN_URL = "https://sei.pe.gov.br/sip/login.php?sigla_orgao_sistema=GOVPE&sigla_sistema=SEI"

XP_USUARIO = '//*[@id="txtUsuario"]'
XP_SENHA = '//*[@id="pwdSenha"]'
CSS_SELECT_ORGAO = '#selOrgao'
XP_BTN_ACESSAR = '//*[@id="sbmAcessar"]'   # ou CSS: '#sbmAcessar'

XP_TXT_PESQUISA_RAPIDA = '//*[@id="txtPesquisaRapida"]'
XP_BTN_LUPA = '//*[@id="spnInfraUnidade"]/img'


def norm(s: str) -> str:
    return " ".join((s or "").strip().lower().split())


def conectar_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAL, scopes)
    return gspread.authorize(creds)


def achar_coluna(headers, *possiveis):
    h_norm = [norm(h) for h in headers]
    for nome in possiveis:
        n = norm(nome)
        if n in h_norm:
            return h_norm.index(n)
    raise KeyError(f"Coluna não encontrada: {possiveis}. Headers: {headers}")


def listar_itens_aguardando(sh):
    """
    Retorna:
      - ws (worksheet)
      - itens: [{"sei": "...", "linha": 2}, ...]
      - idx_status (0-based)
    """
    ws = sh.get_worksheet_by_id(GID_BMS_2026)
    if ws is None:
        raise RuntimeError(f"Não achei worksheet com gid={GID_BMS_2026}")

    valores = ws.get_all_values()
    if not valores or len(valores) < 2:
        return ws, [], None

    headers = valores[0]
    idx_status = achar_coluna(headers, "STATUS", "Status")
    idx_sei = achar_coluna(headers, "N° do SEI", "Nº do SEI", "N° SEI", "Nº SEI")

    itens = []
    seen = set()

    for linha_idx, row in enumerate(valores[1:], start=2):  # start=2 pq linha 1 é header
        if len(row) < len(headers):
            row += [""] * (len(headers) - len(row))

        status = (row[idx_status] or "").strip().upper()
        if status != STATUS_FILTRAR:
            continue

        sei = (row[idx_sei] or "").strip()
        if not sei:
            continue

        # evita duplicados pelo SEI (mantém a primeira ocorrência)
        if sei in seen:
            continue
        seen.add(sei)

        itens.append({"sei": sei, "linha": linha_idx})

    return ws, itens, idx_status


def atualizar_status(ws, linha: int, idx_status: int, novo_status: str):
    # gspread é 1-based para coluna, então +1
    ws.update_cell(linha, idx_status + 1, novo_status)


def sei_quick_search(sb: SB, sei: str) -> None:
    sb.wait_for_element_visible(XP_TXT_PESQUISA_RAPIDA, timeout=60)
    sb.click(XP_TXT_PESQUISA_RAPIDA)
    sb.clear(XP_TXT_PESQUISA_RAPIDA)
    sb.type(XP_TXT_PESQUISA_RAPIDA, sei)
    sb.click(XP_BTN_LUPA)
    sb.sleep(1.2)


def page_or_any_frame_contains(sb: SB, needle: str, timeout: int = 25):
    needle_up = (needle or "").upper()
    end = time.time() + timeout
    last_text = ""

    while time.time() < end:
        try:
            sb.switch_to_default_content()
            txt = sb.execute_script(
                "return (document.body && document.body.innerText) ? document.body.innerText : '';"
            ) or ""
            last_text = txt
            if needle_up in txt.upper():
                return True
        except Exception:
            pass

        try:
            sb.switch_to_default_content()
            frames = sb.find_elements("css selector", "iframe")
        except Exception:
            frames = []

        for fr in frames:
            key = (fr.get_attribute("id") or "").strip() or (fr.get_attribute("name") or "").strip()
            if not key:
                continue
            try:
                sb.switch_to_default_content()
                sb.switch_to_frame(key)
                txt = sb.execute_script(
                    "return (document.body && document.body.innerText) ? document.body.innerText : '';"
                ) or ""
                last_text = txt
                if needle_up in txt.upper():
                    return True
            except Exception:
                continue
            finally:
                try:
                    sb.switch_to_default_content()
                except Exception:
                    pass

        time.sleep(0.5)

    return False


def sei_tem_cehab_gop(sb: SB, timeout: int = 25) -> bool:
    achou = page_or_any_frame_contains(sb, "CEHAB-GOP", timeout=timeout)
    if achou:
        print("   ✅ Achou CEHAB-GOP")
        return True
    print("   ❌ NÃO achou CEHAB-GOP")
    return False


def login_sei(sb: SB, usuario: str, senha: str):
    sb.open(SEI_LOGIN_URL)

    sb.wait_for_element_visible(XP_USUARIO, timeout=60)
    sb.clear(XP_USUARIO)
    sb.type(XP_USUARIO, usuario)

    sb.wait_for_element_visible(XP_SENHA, timeout=60)
    sb.clear(XP_SENHA)
    sb.type(XP_SENHA, senha)

    sb.wait_for_element_visible(CSS_SELECT_ORGAO, timeout=60)
    sb.select_option_by_text(CSS_SELECT_ORGAO, "CEHAB")
    sb.sleep(0.8)

    # garante que o select ficou mesmo em CEHAB
    sb.assert_text("CEHAB", CSS_SELECT_ORGAO)

    sb.wait_for_element_clickable(XP_BTN_ACESSAR, timeout=60)
    sb.click(XP_BTN_ACESSAR)

    # alertas eventuais
    try:
        sb.accept_alert(timeout=3)
    except Exception:
        pass

    # se logou, a pesquisa rápida aparece
    if sb.is_element_visible(XP_TXT_PESQUISA_RAPIDA):
        return

    # fallback: espera um pouco e tenta detectar mensagem de erro
    sb.sleep(2)
    page = (sb.get_text("body") or "").upper()

    if "USUÁRIO" in page or "SENHA" in page or "INVÁLID" in page or "INVALID" in page:
        raise RuntimeError("Falha no login (usuário/senha/órgão). Confira as credenciais e o órgão CEHAB.")

    # última tentativa: esperar mais
    sb.wait_for_element_visible(XP_TXT_PESQUISA_RAPIDA, timeout=90)


def main():
    client = conectar_google_sheets()
    sh = client.open_by_key(PLANILHA_ID)

    ws, itens, idx_status = listar_itens_aguardando(sh)
    print(f"\n✅ SEIs com STATUS '{STATUS_FILTRAR}': {len(itens)}\n")
    if not itens:
        input("👉 ENTER para sair...")
        return

    sei_user = os.getenv("SEI_USER", "marcos.rigel")
    sei_pass = os.getenv("SEI_PASS", "Abc123!@")

    with SB(uc=True, headless=False) as sb:
        sb.maximize_window()
        login_sei(sb, sei_user, sei_pass)

        for i, item in enumerate(itens, start=1):
            sei = item["sei"]
            linha = item["linha"]

            print(f"[{i}/{len(itens)}] 🔎 {sei} (linha {linha})")
            sei_quick_search(sb, sei)

            if sei_tem_cehab_gop(sb, timeout=20):
                # ✅ atualiza planilha
                atualizar_status(ws, linha, idx_status, STATUS_DESTINO)
                print(f"   ✅ STATUS atualizado para '{STATUS_DESTINO}' na linha {linha}")

        print("\n==============================")
        input("👉 Pressione ENTER para fechar o Chrome e encerrar...")


if __name__ == "__main__":
    main()
