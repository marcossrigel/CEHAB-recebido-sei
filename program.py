from oauth2client.service_account import ServiceAccountCredentials
import gspread

from seleniumbase import SB
import os
import time

# ========= Sheets =========
CAMINHO_CREDENCIAL = "formulariosolicitacaopagamento-6292734a5ede.json"
PLANILHA_ID = "1lkM9yOjhu_D2nQjRFl-Wt6lNgWPvzl2wbQiaO633-KM"
GID_BMS_2026 = 1189147903
STATUS_FILTRAR = "AGUARDANDO SEI"

# ========= SEI =========
SEI_LOGIN_URL = "https://sei.pe.gov.br/sip/login.php?sigla_orgao_sistema=GOVPE&sigla_sistema=SEI"

XP_USUARIO = '//*[@id="txtUsuario"]'
XP_SENHA = '//*[@id="pwdSenha"]'
CSS_SELECT_ORGAO = '#selOrgao'
XP_BTN_ACESSAR = '//*[@id="Acessar"]'

XP_TXT_PESQUISA_RAPIDA = '//*[@id="txtPesquisaRapida"]'
XP_BTN_LUPA = '//*[@id="spnInfraUnidade"]/img'


def norm(s: str) -> str:
    return " ".join((s or "").strip().lower().split())


def conectar_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        CAMINHO_CREDENCIAL, scopes
    )
    return gspread.authorize(creds)


def achar_coluna(headers, *possiveis):
    h_norm = [norm(h) for h in headers]
    for nome in possiveis:
        n = norm(nome)
        if n in h_norm:
            return h_norm.index(n)
    raise KeyError(f"Coluna não encontrada: {possiveis}. Headers: {headers}")


def listar_seis_aguardando():
    client = conectar_google_sheets()
    sh = client.open_by_key(PLANILHA_ID)

    ws = sh.get_worksheet_by_id(GID_BMS_2026)
    if ws is None:
        raise RuntimeError(f"Não achei worksheet com gid={GID_BMS_2026}")

    valores = ws.get_all_values()
    if not valores or len(valores) < 2:
        return []

    headers = valores[0]
    idx_status = achar_coluna(headers, "STATUS", "Status")
    idx_sei = achar_coluna(headers, "N° do SEI", "Nº do SEI", "N° SEI", "Nº SEI")

    seis = []
    for row in valores[1:]:
        if len(row) < len(headers):
            row += [""] * (len(headers) - len(row))

        status = (row[idx_status] or "").strip().upper()
        if status == STATUS_FILTRAR:
            sei = (row[idx_sei] or "").strip()
            if sei:
                seis.append(sei)

    # remove duplicados preservando ordem
    seen = set()
    uniq = []
    for s in seis:
        if s not in seen:
            uniq.append(s)
            seen.add(s)
    return uniq


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
        # 1) default_content
        try:
            sb.switch_to_default_content()
            txt = sb.execute_script(
                "return (document.body && document.body.innerText) ? document.body.innerText : '';"
            ) or ""
            last_text = txt
            if needle_up in txt.upper():
                i = txt.upper().find(needle_up)
                return True, "default", txt[max(0, i-80): i+len(needle)+80]
        except Exception:
            pass

        # 2) iframes
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
                    i = txt.upper().find(needle_up)
                    return True, f"iframe:{key}", txt[max(0, i-80): i+len(needle)+80]
            except Exception:
                continue
            finally:
                try:
                    sb.switch_to_default_content()
                except Exception:
                    pass

        time.sleep(0.5)

    return False, "n/a", (last_text[:300] if last_text else "")


def sei_tem_cehab_gop(sb: SB, timeout: int = 25) -> bool:
    achou, onde, _trecho = page_or_any_frame_contains(sb, "CEHAB-GOP", timeout=timeout)
    if achou:
        print(f"   ✅ Achou CEHAB-GOP em: {onde}")
        return True
    print("   ❌ NÃO achou CEHAB-GOP no texto visível (default/iframes).")
    return False


def login_sei(sb: SB, usuario: str, senha: str):
    sb.open(SEI_LOGIN_URL)
    sb.wait_for_element_visible(XP_USUARIO, timeout=60)
    sb.type(XP_USUARIO, usuario)

    sb.wait_for_element_visible(XP_SENHA, timeout=60)
    sb.type(XP_SENHA, senha)

    sb.wait_for_element_visible(CSS_SELECT_ORGAO, timeout=60)
    sb.select_option_by_text(CSS_SELECT_ORGAO, "CEHAB")
    sb.sleep(0.3)

    sb.wait_for_element_visible(XP_BTN_ACESSAR, timeout=60)
    sb.click(XP_BTN_ACESSAR)
    sb.sleep(1.2)

    # às vezes aparece alert
    try:
        sb.accept_alert(timeout=2)
    except Exception:
        pass

    # garante que chegou na tela principal
    sb.wait_for_element_visible(XP_TXT_PESQUISA_RAPIDA, timeout=90)


def main():
    seis = listar_seis_aguardando()
    print(f"\n✅ SEIs com STATUS '{STATUS_FILTRAR}': {len(seis)}\n")

    if not seis:
        return

    sei_user = os.getenv("SEI_USER", "marcos.rigel")
    sei_pass = os.getenv("SEI_PASS", "Abc123!@")

    seis_com_cehab_gop = []

    with SB(uc=False, headless=False) as sb:
        sb.maximize_window()
        login_sei(sb, sei_user, sei_pass)

        for i, sei in enumerate(seis, start=1):
            print(f"[{i}/{len(seis)}] 🔎 {sei}")
            sei_quick_search(sb, sei)

            # aqui é o “quadro branco”
            if sei_tem_cehab_gop(sb, timeout=20):
                seis_com_cehab_gop.append(sei)
        print("\n==============================")
        input("👉 Pressione ENTER para fechar o Chrome...")

if __name__ == "__main__":
    main()
