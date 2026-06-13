"""
Automação para coleta de descrições e imagens de produtos farmacêuticos.
Sites: Drogaria São Paulo (API VTEX) e Panvel (SPA Angular - usa Playwright)
"""

import os
import re
import json
import time
import hashlib
import requests
import openpyxl
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

# ─── CONFIGURAÇÕES ────────────────────────────────────────────────────────────
PLANILHA_ENTRADA   = "produtos.xlsx"  # <-- nome da sua planilha
COLUNA_EAN         = "A"             # coluna com os EANs
COLUNA_DESC_DSP    = "B"             # coluna Descrição Drogaria SP
COLUNA_DESC_PANVEL = "C"             # coluna Descrição Panvel
PASTA_IMAGENS      = "imagens"       # pasta raiz para salvar imagens
MAX_IMAGENS        = 10              # limite de imagens únicas por EAN

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9",
}

# ─── UTILITÁRIOS ──────────────────────────────────────────────────────────────

def hash_imagem(conteudo: bytes) -> str:
    return hashlib.md5(conteudo).hexdigest()


def baixar_imagem(url: str, session: requests.Session) -> bytes | None:
    try:
        r = session.get(url, timeout=15, headers=HEADERS)
        if r.status_code == 200 and "image" in r.headers.get("content-type", ""):
            return r.content
    except Exception as e:
        print(f"    ✗ Erro ao baixar {url}: {e}")
    return None


def salvar_imagens_unicas(ean: str, urls: list, session: requests.Session) -> int:
    """
    Baixa e salva até MAX_IMAGENS imagens únicas (por hash MD5).
    Pasta: {PASTA_IMAGENS}/{EAN}/
    Nomes: EAN - 1.jpg, EAN - 2.jpg, ...
    """
    pasta_ean = Path(PASTA_IMAGENS) / str(ean)
    pasta_ean.mkdir(parents=True, exist_ok=True)

    hashes_vistos: set = set()
    contador = 1

    for url in urls:
        if contador > MAX_IMAGENS:
            break
        conteudo = baixar_imagem(url, session)
        if conteudo is None:
            continue
        h = hash_imagem(conteudo)
        if h in hashes_vistos:
            print(f"    ⚠ Duplicada ignorada: {url}")
            continue
        hashes_vistos.add(h)

        ext = url.split("?")[0].rsplit(".", 1)[-1].lower()
        ext = ext if ext in {"jpg", "jpeg", "png", "webp", "gif"} else "jpg"

        nome = pasta_ean / f"{ean} - {contador}.{ext}"
        nome.write_bytes(conteudo)
        print(f"    ✓ Salva: {nome.name}")
        contador += 1

    return contador - 1


def col_letter_to_index(letter: str) -> int:
    letter = letter.upper().strip()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


# ─── DROGARIA SÃO PAULO (API VTEX) ────────────────────────────────────────────

DSP_BASE = "https://www.drogariasaopaulo.com.br"

def coletar_dsp(ean: str, session: requests.Session) -> tuple:
    """Usa a API VTEX nativa — rápido, sem JavaScript."""
    print(f"  [DSP] Buscando EAN {ean}...")
    url = (
        f"{DSP_BASE}/api/catalog_system/pub/products/search"
        f"?fq=alternateIds_Ean:{ean}&_from=0&_to=0"
    )
    try:
        r = session.get(url, headers=HEADERS, timeout=20)
        data = r.json()
    except Exception as e:
        print(f"  ✗ DSP: Erro na API: {e}")
        return "", []

    if not data:
        print(f"  – DSP: Não encontrado para EAN {ean}")
        return "", []

    produto = data[0]
    todas_imgs = []
    for sku in produto.get("items", []):
        for img in sku.get("images", []):
            img_url = img.get("imageUrl", "")
            if img_url and img_url not in todas_imgs:
                todas_imgs.append(img_url)

    descricao = produto.get("description") or produto.get("productName", "")
    print(f"  [DSP] ✓ {produto.get('productName', '')} | {len(todas_imgs)} imagens")
    return descricao, todas_imgs


# ─── PANVEL (Playwright — SPA Angular) ────────────────────────────────────────

PANVEL_BUSCA = "https://www.panvel.com/panvel/buscarProduto.do?termoPesquisa={ean}"

def coletar_panvel(ean: str, page) -> tuple:
    """
    Usa Playwright para renderizar o JavaScript da Panvel.
    Correção: wait_for_selector em script[type=ld+json] falha pois
    <script> tags têm display:none e nunca são 'visíveis' para o Playwright.
    Solução: esperar pelo h1.product-title (visível) e ler o JSON-LD depois.
    """
    print(f"  [Panvel] Buscando EAN {ean}...")
    try:
        # ── 1. Página de busca ─────────────────────────────────────────────
        page.goto(
            f"https://www.panvel.com/panvel/buscarProduto.do?termoPesquisa={ean}",
            wait_until="domcontentloaded",
            timeout=30000
        )

        # Aguarda o Angular renderizar os links de produto
        try:
            page.wait_for_selector('a[href*="/p-"]', state="visible", timeout=15000)
        except PwTimeout:
            print(f"  – Panvel: Produto não encontrado para EAN {ean}")
            return "", []

        # ── 2. Pega o link do primeiro produto ─────────────────────────────
        produto_url = page.locator('a[href*="/p-"]').first.get_attribute("href")
        if not produto_url:
            print(f"  – Panvel: Link não encontrado para EAN {ean}")
            return "", []

        if produto_url.startswith("/"):
            produto_url = "https://www.panvel.com" + produto_url

        # ── 3. Abre a página do produto ────────────────────────────────────
        time.sleep(0.5)
        page.goto(produto_url, wait_until="domcontentloaded", timeout=30000)

        # Espera o h1 do produto ficar visível (confirma que a página carregou)
        # CORREÇÃO: não esperar script[type="application/ld+json"] — tem display:none
        try:
            page.wait_for_selector("h1.product-title", state="visible", timeout=15000)
        except PwTimeout:
            # fallback: esperar qualquer h1
            try:
                page.wait_for_selector("h1", state="visible", timeout=10000)
            except PwTimeout:
                print(f"  ⚠ Panvel: Página do produto demorou para carregar ({produto_url})")

        # ── 4. Lê o JSON-LD diretamente do DOM (state="attached" = só precisa existir) ──
        # <script> tags existem no DOM mas têm display:none — usar evaluate, não wait_for_selector
        ld_json_list = page.evaluate("""() => {
            const scripts = document.querySelectorAll('script[type="application/ld+json"]');
            return Array.from(scripts).map(el => {
                try { return JSON.parse(el.textContent); } catch(e) { return null; }
            }).filter(Boolean);
        }""")

        produto_ld = next((d for d in ld_json_list if d.get("@type") == "Product"), None)

        if produto_ld:
            images = produto_ld.get("image", [])
            if isinstance(images, str):
                images = [images]
            desc = produto_ld.get("description", "")
            nome = produto_ld.get("name", "")
            print(f"  [Panvel] ✓ {nome} | {len(images)} imagens")
            return desc, images

        # ── 5. Fallback: extrair do DOM caso JSON-LD não esteja disponível ─
        titulo = page.locator("h1").first.text_content(timeout=5000).strip()
        imgs = page.evaluate("""() =>
            [...new Set(
                Array.from(document.querySelectorAll("img[src*='staticpanvel']"))
                    .map(e => e.src)
            )]
        """)
        desc_items = page.locator(".destaques-produto li").all_text_contents()
        desc = " | ".join(i.strip() for i in desc_items if i.strip())
        print(f"  [Panvel] ✓ (fallback) {titulo} | {len(imgs)} imagens")
        return desc, imgs

    except Exception as e:
        print(f"  ✗ Panvel: Erro para EAN {ean}: {e}")
        return "", []


# ─── PROCESSAMENTO PRINCIPAL ──────────────────────────────────────────────────

def processar_planilha():
    print("=" * 60)
    print("  Automação de Coleta - Farmácias Online")
    print("=" * 60)

    if not Path(PLANILHA_ENTRADA).exists():
        print(f"✗ Planilha '{PLANILHA_ENTRADA}' não encontrada!")
        return

    wb  = openpyxl.load_workbook(PLANILHA_ENTRADA)
    ws  = wb.active

    col_ean    = col_letter_to_index(COLUNA_EAN)
    col_dsp    = col_letter_to_index(COLUNA_DESC_DSP)
    col_panvel = col_letter_to_index(COLUNA_DESC_PANVEL)

    # Cabeçalhos (se vazios)
    if not ws.cell(row=1, column=col_dsp).value:
        ws.cell(row=1, column=col_dsp).value    = "Descrição Drogaria SP"
    if not ws.cell(row=1, column=col_panvel).value:
        ws.cell(row=1, column=col_panvel).value  = "Descrição Panvel"

    # Coleta os EANs
    eans = []
    for row in ws.iter_rows(min_row=2):
        cell = row[col_ean - 1]
        if cell.value is not None:
            eans.append((cell.row, str(cell.value).strip()))

    print(f"\n📋 {len(eans)} EAN(s) encontrado(s).\n")

    http_session = requests.Session()
    http_session.headers.update(HEADERS)

    # Inicia o Playwright (um browser para todos os EANs — mais eficiente)
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)  # mude para False para ver o browser
        context = browser.new_context(
            user_agent=HEADERS["User-Agent"],
            locale="pt-BR",
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()

        for idx, (row_num, ean) in enumerate(eans, 1):
            print(f"\n[{idx}/{len(eans)}] EAN: {ean}")
            print("-" * 50)

            # ── Drogaria São Paulo ─────────────────────────────────────────
            desc_dsp, imgs_dsp = coletar_dsp(ean, http_session)
            ws.cell(row=row_num, column=col_dsp).value = desc_dsp

            # ── Panvel ────────────────────────────────────────────────────
            desc_panvel, imgs_panvel = coletar_panvel(ean, page)
            ws.cell(row=row_num, column=col_panvel).value = desc_panvel

            # ── Imagens (DSP + Panvel, sem duplicatas, max 10) ────────────
            todas_urls = list(dict.fromkeys(imgs_dsp + imgs_panvel))  # deduplicar por URL
            if todas_urls:
                print(f"\n  📥 Baixando imagens ({len(todas_urls)} URLs únicas)...")
                total = salvar_imagens_unicas(ean, todas_urls, http_session)
                print(f"  ✅ {total} imagem(ns) salva(s) em '{PASTA_IMAGENS}/{ean}/'")
            else:
                print(f"  ⚠ Nenhuma imagem encontrada para EAN {ean}")

            # Salva a planilha após cada EAN (segurança)
            wb.save(PLANILHA_ENTRADA)
            time.sleep(1.5)

        browser.close()

    print("\n" + "=" * 60)
    print(f"✅ Concluído! Planilha salva: {PLANILHA_ENTRADA}")
    print(f"📁 Imagens em: {Path(PASTA_IMAGENS).resolve()}")
    print("=" * 60)


if __name__ == "__main__":
    processar_planilha()
    