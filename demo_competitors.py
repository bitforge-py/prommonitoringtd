"""Порівняння цін вітрини клієнта на Prom з мінімальними цінами конкурентів.

Крок 1. З TARGET_SHOP_URL збираються товари: назва, модель, ціна в грн, посилання.
Крок 2. Для кожного товару виконується пошук на prom.ua з сортуванням за зростанням
        ціни. Відсіюються аксесуари, власний магазин клієнта і товари не в наявності.
Крок 3. Звіт зберігається в Excel. Різниця цін підсвічується кольором.

Перед запуском замініть TARGET_SHOP_URL на вітрину клієнта.
Запуск: python demo_competitors.py
"""
from __future__ import annotations

import asyncio
import random
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import quote_plus, urlsplit

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.async_api import (
    BrowserContext,
    Page,
    Route,
    async_playwright,
)
from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
)
from rich.table import Table

# ---------------------------------------------------------------------------
# Налаштування. Перед запуском замініть приклад на вітрину клієнта.
# ---------------------------------------------------------------------------
# Сайт або вітрина клієнта на Prom.
TARGET_SHOP_URL = "https://example-shop.prom.ua"
# Скільки товарів клієнта взяти в звіт.
MAX_PRODUCTS = 15
OUTPUT_FILE = "demo_price_comparison.xlsx"

# Необов'язковий розділ каталогу на тому самому домені, що й TARGET_SHOP_URL.
# Порожній рядок означає весь каталог (/ua/product_list).
CATALOG_URL = ""

# Якщо ціна конкурента менша за цю частку від ціни клієнта, вважаємо, що це
# аксесуар/кабель/чохол, а не той самий товар, і пропускаємо.
MIN_PRICE_RATIO = 0.40

# --- Браузер ---------------------------------------------------------------
HEADLESS = True
LOCALE = "uk-UA"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
VIEWPORT = {"width": 1366, "height": 900}
BLOCKED_RESOURCE_TYPES = {"image", "font", "media"}
NAV_TIMEOUT_MS = 30_000
ACTION_TIMEOUT_MS = 8_000
MIN_DELAY, MAX_DELAY = 1.5, 3.5          # пауза між запитами до Prom, сек
MAX_CATALOG_PAGES = 10                   # максимум сторінок каталогу клієнта
SEARCH_URL = "https://prom.ua/ua/search?search_term={query}&sort=price"

console = Console()

HEADERS = [
    "Товар клиента",
    "Цена клиента (грн)",
    "Мин. цена конкурента (грн)",
    "Разница (грн)",
    "Конкурент",
    "Ссылка на конкурента",
    "Ссылка на товар клиента",
]
COL_WIDTH_LIMITS = {1: 55, 5: 35, 6: 60, 7: 60}
DEFAULT_COL_WIDTH_LIMIT = 28

FILL_WORSE = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")   # клієнт дорожчий
FILL_BETTER = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")  # клієнт дешевший
FILL_HEADER = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# Моделі
# ---------------------------------------------------------------------------
@dataclass
class ClientProduct:
    name: str
    price: float
    url: str
    models: list[str] = field(default_factory=list)   # кандидати в модель/артикул
    sku: str = ""
    company_id: str = ""
    company_name: str = ""

    @property
    def model(self) -> str:
        return self.models[0] if self.models else ""


@dataclass
class Competitor:
    price: float
    shop: str
    url: str
    title: str = ""


@dataclass
class Result:
    product: ClientProduct
    competitor: Competitor | None = None
    query: str = ""
    note: str = ""


# ---------------------------------------------------------------------------
# Утиліти: ціни, назви, моделі, пошукові запити
# ---------------------------------------------------------------------------
PRICE_RE = re.compile(r"\d[\d\s\u00a0\u202f]*(?:[.,]\d+)?")


def parse_price(text: str | None) -> float | None:
    """'3 019 ₴' -> 3019.0. Для рядка з двома цінами береться перша."""
    if not text:
        return None
    match = PRICE_RE.search(text)
    if not match:
        return None
    raw = re.sub(r"[\s\u00a0\u202f]", "", match.group(0)).replace(",", ".")
    try:
        value = float(raw)
    except ValueError:
        return None
    return value if value > 0 else None


def clean_text(value: str | None, limit: int = 300) -> str:
    return re.sub(r"\s+", " ", value or "").strip()[:limit]


# Маркетингові слова, які не повинні потрапити в пошуковий запит.
MARKETING_RE = re.compile(
    r"\b(?:купити|купить|купуй\w*|акці[яїюй]\w*|акци[яиюй]\w*|знижк\w*|скидк\w*|"
    r"розпродаж\w*|распродаж\w*|хіт|хит|новинк\w*|безкоштовн\w*|бесплатн\w*|"
    r"доставк\w*|оригінал\w*|оригинал\w*|в\s+наявності|в\s+наличии|топ|top|sale|"
    r"ціна|цена|недорого|дешево|офіційн\w*|официальн\w*|найкращ\w*|лучш\w*|"
    r"гаранті\w*(?:\s+\d+\s*(?:років|роки|рік|міс\w*|года?|лет|мес\w*))?|"
    r"гарантия(?:\s+\d+\s*\w+)?)\b",
    re.IGNORECASE,
)
NOISE_CHARS_RE = re.compile(r"[|!?•«»\"“”„()\[\],;:/\\№#*~_]+")
CODE_IN_BRACKETS_RE = re.compile(r"[(\[]\s*\d{6,}\s*[)\]]")   # внутрішні коди типу (00000099417)

# Одиниці виміру/характеристики, які не є моделлю (30000mAh, 10W, 5G, 4K, 3in1 ...).
SPEC_TOKEN_RE = re.compile(
    r"^\d+(?:[.,]\d+)?(?:mah|мач|w|вт|wh|вт\.?год|v|в|a|а|gb|гб|tb|тб|mb|мб|mhz|ггц|ghz|hz|гц|"
    r"mm|мм|cm|см|m|м|kg|кг|g|г|l|л|ml|мл|mp|мп|k|fps|ms|db|дб|x|in\d+|rpm|ом|ohm|dpi|nm|"
    r"pcs|шт|pack|usb|inch|дюйм\w*)$",
    re.IGNORECASE,
)
UNIT_ONLY_RE = re.compile(r"^(?:gb|гб|tb|тб|mb|мб|v|в|w|вт|wh|mah|a|а|hz|гц|mm|мм|cm|см|kg|кг)$", re.IGNORECASE)
# Слова-варіанти: «DELTA 2» і «DELTA 2 Max» - різні товари.
VARIANT_WORDS = {
    "max", "pro", "plus", "ultra", "mini", "lite", "neo", "se", "air", "fe", "xl", "extra",
    "макс", "про", "плюс", "ультра", "міні",
}
# Стандарти з короткою цифрою (USB3.0, DDR5, QC3.0, IP67). Не плутати з AC180P, AC70P.
STANDARD_TOKEN_RE = re.compile(
    r"^(?:usb|type|hdmi|wi-?fi|bt|bluetooth|ddr|lpddr|pcie|sata|nvme|ips|oled|amoled|lcd|led|"
    r"qc|pd|tws|rgb|ram|rom|ssd|hdd|fhd|uhd|hd|ip)[-]?\d{1,2}(?:\.\d)?$",
    re.IGNORECASE,
)
# Ємність (Wh, mAh): різні значення означають різні товари (DELTA 3 1024Wh != DELTA 3 Air 960Wh).
CAPACITY_RE = re.compile(
    r"(\d{1,3}(?:[\s\u00a0]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?)\s*"
    r"(wh|вт\s*[·\-.*/]?\s*(?:год|г)(?![а-яіїє])|mah|мач)",
    re.IGNORECASE,
)
COLOR_WORDS = {
    "black", "white", "gray", "grey", "silver", "gold", "blue", "red", "green", "pink",
    "orange", "yellow", "purple", "titanium", "midnight", "starlight",
    "чорний", "білий", "сірий", "синій", "червоний", "зелений", "рожевий", "золотий",
}
PREPOSITIONS_FOR = {"для", "for", "под", "до", "сумісний", "сумісна", "compatible"}

# Слова, що вказують на аксесуар. Враховуються, лише якщо їх немає в назві товару клієнта.
ACCESSORY_RE = re.compile(
    r"чохол|чехол|\bкабел|\bшнур|скло|стекло|плівк|пленк|захисн|бампер|кейс|\bcase\b|\bcover\b|"
    r"\bсумк|рюкзак|перехідник|переходник|тримач|держател|кріплення|крепление|наклейк|розетк|"
    r"комплектуюч|запчаст|ремінець|ремешок|мішок|фільтр\b",
    re.IGNORECASE,
)

IN_STOCK_RE = re.compile(r"в\s+наявності|в\s+наличии|готово\s+до\s+відправки|готов\s+к\s+отправке", re.I)


def clean_query_text(name: str) -> str:
    text = CODE_IN_BRACKETS_RE.sub(" ", name)
    text = MARKETING_RE.sub(" ", text)
    text = NOISE_CHARS_RE.sub(" ", text)
    tokens = [t for t in text.split() if re.search(r"\w", t)]
    return " ".join(tokens)


def _token_norm(token: str) -> str:
    return re.sub(r"[^0-9a-zа-яіїєґ]", "", token.lower())


def split_tokens(text: str) -> list[str]:
    return [t for t in re.split(r"[\s/,;()\[\]|+]+", text) if t]


def looks_like_model(token: str) -> int:
    """0 - не модель, інакше пріоритет (більше = краще)."""
    core = token.strip(".,-")
    norm = _token_norm(core)
    if len(norm) < 3 or core.lower() in COLOR_WORDS:
        return 0
    if SPEC_TOKEN_RE.match(core) or STANDARD_TOKEN_RE.match(core):
        return 0
    has_letter = bool(re.search(r"[a-zа-яіїєґ]", norm))
    has_digit = bool(re.search(r"\d", norm))
    if has_letter and has_digit:
        return 3 if "-" in core else 2
    if has_digit and not has_letter and 5 <= len(norm) <= 10 and not re.fullmatch(r"(?:19|20)\d\d", norm):
        return 1   # чисто цифровий код моделі (44998, 930487)
    return 0


def extract_models(name: str, limit: int = 3) -> list[str]:
    """Кандидати в модель/артикул за спаданням надійності.

    1) токени з літерами й цифрами (TX-30, PB930487, J101B, EFDELTA2-EU);
    2) чисто цифрові коди (44998);
    3) пари «латинське слово + число» (DELTA 2, Explorer 1000, iPhone 15).
    """
    tokens = split_tokens(clean_query_text(name))
    scored: list[tuple[int, int, str]] = []   # (пріоритет, порядок, модель)
    for index, token in enumerate(tokens):
        score = looks_like_model(token)
        if score:
            scored.append((score, -index, token.strip(".,-")))
    # Пари «слово + число» слабкий ознака, тому лише коли точного коду в назві немає
    for index in range(len(tokens) - 1 if not scored else 0):
        word, number = tokens[index], tokens[index + 1]
        next_token = tokens[index + 2] if index + 2 < len(tokens) else ""
        if (
            re.fullmatch(r"[A-Za-z]{3,}", word)
            and word.lower() not in COLOR_WORDS
            and re.fullmatch(r"\d{1,4}", number)
            and not UNIT_ONLY_RE.match(next_token)
        ):
            parts = [word, number]
            # «DELTA 3 Max Plus» - варіанти входять до складу моделі
            for variant in tokens[index + 2 : index + 4]:
                if variant.lower() not in VARIANT_WORDS:
                    break
                parts.append(variant)
            scored.append((1, -index, " ".join(parts)))
    scored.sort(key=lambda item: (-item[0], -item[1]))
    result: list[str] = []
    for _, _, model in scored:
        if model.lower() not in {m.lower() for m in result}:
            result.append(model)
    return result[:limit]


def guess_brand(name: str, model: str) -> str:
    """Слово, що стоїть перед моделлю (Hoco J101B -> Hoco), якщо це схоже на бренд."""
    tokens = split_tokens(clean_query_text(name))
    index = find_model(model, " ".join(tokens))[0]
    if index > 0:
        prev = tokens[index - 1]
        if (
            re.fullmatch(r"[A-Za-zА-Яа-яІіЇїЄєҐґ][\w\-]{1,}", prev)
            and prev.lower() not in COLOR_WORDS
            and prev.lower() not in VARIANT_WORDS
        ):
            return prev
    return ""


MAX_QUERIES = 4


def build_queries(name: str, models: list[str]) -> list[str]:
    """Драбина запитів від точного до широкого: повна назва -> бренд+модель -> модель."""
    words = clean_query_text(name).split()
    queries = [" ".join(words[:12])]
    for model in models:
        brand = guess_brand(name, model)
        queries.append(f"{brand} {model}".strip())
    if models:
        queries.append(models[0])
    seen: set[str] = set()
    result: list[str] = []
    for query in queries:
        key = query.lower()
        if query and key not in seen:
            seen.add(key)
            result.append(query)
    return result[:MAX_QUERIES]


def find_model(model: str, title: str) -> tuple[int, int]:
    """(індекс першого токена, кількість токенів) або (-1, 0)."""
    target = _token_norm(model)
    if not target:
        return -1, 0
    norms = [_token_norm(t) for t in split_tokens(title)]
    for size in (1, 2, 3):
        for index in range(len(norms) - size + 1):
            if "".join(norms[index : index + size]) == target:
                return index, size
    return -1, 0


def significant_tokens(text: str) -> set[str]:
    return {
        _token_norm(t)
        for t in split_tokens(text)
        if len(_token_norm(t)) >= 3 and _token_norm(t) not in COLOR_WORDS
    }


def extract_capacities(text: str) -> dict[str, set[float]]:
    """{'wh': {1024.0}, 'mah': {30000.0}} з назви товару."""
    result: dict[str, set[float]] = {}
    for match in CAPACITY_RE.finditer(text):
        raw = re.sub(r"[\s\u00a0]", "", match.group(1)).replace(",", ".")
        kind = "mah" if match.group(2).lower() in {"mah", "мач"} else "wh"
        try:
            result.setdefault(kind, set()).add(float(raw))
        except ValueError:
            continue
    return result


def capacity_conflict(client_name: str, title: str) -> bool:
    """True, якщо в обох назвах вказана ємність одного типу, але значення різні."""
    mine, theirs = extract_capacities(client_name), extract_capacities(title)
    for kind, values in mine.items():
        other = theirs.get(kind)
        if other and not any(abs(a - b) <= max(a, b) * 0.01 for a in values for b in other):
            return True
    return False


def is_relevant(product: ClientProduct, title: str, partial: bool) -> bool:
    """Чи це той самий товар, а не схожий чи аксесуар до нього."""
    client_lower = product.name.lower()
    if ACCESSORY_RE.search(title):
        accessory_words = set(m.group(0) for m in ACCESSORY_RE.finditer(title.lower()))
        if any(word.strip() not in client_lower for word in accessory_words):
            return False
    if capacity_conflict(product.name, title):
        return False

    if product.models:
        title_tokens = split_tokens(title)
        client_tokens = {t.lower() for t in split_tokens(product.name)}
        for model in product.models:
            position, size = find_model(model, title)
            if position < 0:
                continue
            # «Чохол ДЛЯ EcoFlow DELTA 2» - це аксесуар до моделі, а не сама модель
            before = {t.lower() for t in title_tokens[:position]}
            if (before & PREPOSITIONS_FOR) - client_tokens:
                continue
            # «DELTA 2» і «DELTA 2 Max» - різні товари
            end = position + size
            follower = title_tokens[end].lower() if end < len(title_tokens) else ""
            if follower in VARIANT_WORDS and follower not in client_tokens:
                continue
            return True
        return False

    # Моделі немає: порівнюємо за перетином значущих слів і ігноруємо часткові збіги Prom
    if partial:
        return False
    mine = significant_tokens(clean_query_text(product.name))
    theirs = significant_tokens(title)
    return bool(mine) and len(mine & theirs) / len(mine) >= 0.7


# ---------------------------------------------------------------------------
# Браузер
# ---------------------------------------------------------------------------
async def block_heavy_resources(route: Route) -> None:
    if route.request.resource_type in BLOCKED_RESOURCE_TYPES:
        await route.abort()
    else:
        await route.continue_()


async def new_context(playwright) -> tuple[object, BrowserContext]:
    browser = await playwright.chromium.launch(
        headless=HEADLESS,
        args=["--disable-blink-features=AutomationControlled"],
    )
    context = await browser.new_context(
        user_agent=USER_AGENT,
        locale=LOCALE,
        viewport=VIEWPORT,
    )
    context.set_default_timeout(ACTION_TIMEOUT_MS)
    context.set_default_navigation_timeout(NAV_TIMEOUT_MS)
    await context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    await context.route("**/*", block_heavy_resources)
    return browser, context


async def pause(min_s: float = MIN_DELAY, max_s: float = MAX_DELAY) -> None:
    await asyncio.sleep(random.uniform(min_s, max_s))


async def scroll_a_bit(page: Page, steps: int = 4) -> None:
    """Невелика прокрутка, щоб підвантажилися відкладені картки."""
    for _ in range(steps):
        await page.mouse.wheel(0, random.randint(600, 900))
        await asyncio.sleep(random.uniform(0.3, 0.6))


# ---------------------------------------------------------------------------
# Крок 1: товари клієнта
# ---------------------------------------------------------------------------
CLIENT_TILES_JS = """
() => {
  const text = (root, sel) => { const e = root.querySelector(sel); return e ? (e.innerText || e.textContent || '').trim() : ''; };
  const company = document.querySelector('[data-qaid="company_name"]');
  const blocks = [...document.querySelectorAll('[data-qaid="product-block"]')];
  return {
    companyName: company ? (company.innerText || '').trim() : '',
    items: blocks.map(b => {
      const link = b.querySelector('a.b-goods-title') || b.querySelector('a[href*="/p"]');
      const buy = b.querySelector('[data-qaid="buy-button"]');
      return {
        name: link ? (link.getAttribute('title') || link.innerText || '').trim() : '',
        href: link ? link.href : '',
        price: text(b, '.b-goods-price__value_type_current') || text(b, '.b-goods-price') ||
               (buy ? (buy.getAttribute('data-product-price') || '') : ''),
        sku: text(b, '.b-goods-sku'),
        companyId: buy ? (buy.getAttribute('data-company-id') || '') : '',
        presence: text(b, '[data-qaid="presence_data"]')
      };
    })
  };
}
"""


def catalog_start_url() -> str:
    base = TARGET_SHOP_URL.rstrip("/")
    if CATALOG_URL:
        if urlsplit(CATALOG_URL).hostname == urlsplit(TARGET_SHOP_URL).hostname:
            return CATALOG_URL.rstrip("/")
        console.print(
            "[yellow]CATALOG_URL з іншого домену, ніж TARGET_SHOP_URL. Ігнорую, беру весь каталог.[/yellow]"
        )
    return f"{base}/ua/product_list"


def catalog_page_url(start: str, number: int) -> str:
    return start if number == 1 else f"{start}/page_{number}"


async def collect_client_products(page: Page) -> list[ClientProduct]:
    console.rule("[bold cyan]Крок 1: збір товарів клієнта")
    start = catalog_start_url()
    products: list[ClientProduct] = []
    seen_urls: set[str] = set()
    for number in range(1, MAX_CATALOG_PAGES + 1):
        url = catalog_page_url(start, number)
        try:
            response = await page.goto(url, wait_until="domcontentloaded")
        except Exception as exc:  # noqa: BLE001
            console.print(f"[red]{url}: {type(exc).__name__}: {exc}[/red]")
            break
        if response and response.status >= 400:
            console.print(f"[yellow]{url}: HTTP {response.status}[/yellow]")
            break
        try:
            await page.wait_for_selector('[data-qaid="product-block"]', timeout=10_000)
        except Exception:  # noqa: BLE001
            pass
        data = await page.evaluate(CLIENT_TILES_JS)
        added = 0
        for item in data["items"]:
            product_url = item["href"].split("?")[0].split("#")[0]
            name = clean_text(item["name"])
            price = parse_price(item["price"])
            if not name or not product_url or product_url in seen_urls:
                continue
            seen_urls.add(product_url)
            if price is None:
                console.print(f"[dim]SKIP[/dim] {name[:60]}: немає ціни")
                continue
            sku_match = re.search(r"[\w\-]+", item["sku"].replace("Код:", ""))
            products.append(
                ClientProduct(
                    name=name,
                    price=price,
                    url=product_url,
                    models=extract_models(name),
                    sku=sku_match.group(0) if sku_match else "",
                    company_id=item["companyId"],
                    company_name=re.sub(r"^Продавець\s+", "", data["companyName"]).strip(),
                )
            )
            added += 1
            if len(products) >= MAX_PRODUCTS:
                break
        console.print(f"Сторінка {number}: додано [green]{added}[/green], всього [bold]{len(products)}[/bold]")
        if len(products) >= MAX_PRODUCTS or not data["items"] or added == 0:
            break
        await pause(1.0, 2.0)
    return products[:MAX_PRODUCTS]


# ---------------------------------------------------------------------------
# Крок 2: пошук конкурентів на Prom
# ---------------------------------------------------------------------------
SEARCH_BLOCKS_JS = """
() => {
  const marker = document.querySelector('[data-qaid="partial_match_products"]');
  const text = (root, sel) => { const e = root.querySelector(sel); return e ? (e.innerText || '').trim() : ''; };
  return [...document.querySelectorAll('[data-qaid="product_block"]')].map(b => {
    const priceEl = b.querySelector('[data-qaid="product_price"]');
    const link = b.querySelector('[data-qaid="product_link"]') || b.querySelector('a[href*="prom.ua"], a[href^="/ua/"]');
    const nameEl = b.querySelector('[data-qaid="product_name"]');
    return {
      name: nameEl ? (nameEl.innerText || '').trim() : '',
      priceAttr: priceEl ? (priceEl.getAttribute('data-qaprice') || '') : '',
      priceText: priceEl ? (priceEl.innerText || '').trim() : '',
      presence: text(b, '[data-qaid="product_presence"]'),
      company: text(b, '[data-qaid="company_name"]'),
      companyId: b.getAttribute('data-company-id') || '',
      href: link ? link.href : '',
      partial: marker ? (marker.contains(b) || !!(marker.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING)) : false
    };
  });
}
"""

CAPTCHA_RE = re.compile(r"captcha|капч|are you a robot|перевірка безпеки|доступ обмежено", re.I)


async def fetch_search_blocks(page: Page, query: str) -> list[dict]:
    url = SEARCH_URL.format(query=quote_plus(query))
    response = await page.goto(url, wait_until="domcontentloaded")
    if response and response.status >= 400:
        raise RuntimeError(f"HTTP {response.status}")
    try:
        await page.wait_for_selector('[data-qaid="product_block"]', timeout=10_000)
    except Exception:  # noqa: BLE001
        title = await page.title()
        if CAPTCHA_RE.search(title):
            raise RuntimeError("Prom показав капчу")
        return []
    await scroll_a_bit(page)
    return await page.evaluate(SEARCH_BLOCKS_JS)


def normalize_shop_name(name: str) -> str:
    return re.sub(r"[^0-9a-zа-яіїєґ]", "", re.sub(r"^Продавець\s+", "", name, flags=re.I).lower())


def pick_best_competitor(product: ClientProduct, blocks: list[dict]) -> Competitor | None:
    own_name = normalize_shop_name(product.company_name)
    floor = product.price * MIN_PRICE_RATIO
    best: Competitor | None = None
    for block in blocks:
        price = parse_price(block["priceAttr"]) or parse_price(block["priceText"])
        title = clean_text(block["name"])
        shop = re.sub(r"^Продавець\s+", "", clean_text(block["company"], 120), flags=re.I)
        if price is None or not title or not block["href"]:
            continue
        # власний магазин клієнта
        if product.company_id and block["companyId"] == product.company_id:
            continue
        if own_name and normalize_shop_name(shop) == own_name:
            continue
        # лише ті, що є в наявності
        if not IN_STOCK_RE.search(block["presence"] or ""):
            continue
        # дешевше 40% від ціни клієнта - це аксесуар, а не такий самий товар
        if price < floor:
            continue
        if not is_relevant(product, title, block["partial"]):
            continue
        if best is None or price < best.price:
            best = Competitor(price=price, shop=shop or "Невідомий магазин", url=block["href"], title=title)
    return best


async def find_competitor(page: Page, product: ClientProduct) -> tuple[Competitor | None, str]:
    """Драбина запитів: повна назва -> бренд+модель -> модель.

    Перший запит часто знаходить лише одну-дві ідентичні пропозиції, тому після
    першого успіху виконується ще один (ширший) запит і береться мінімум із двох.
    Далі драбина зупиняється.
    """
    best: Competitor | None = None
    best_query = ""
    queries = build_queries(product.name, product.models)
    for index, query in enumerate(queries):
        blocks = await fetch_search_blocks(page, query)
        competitor = pick_best_competitor(product, blocks)
        if competitor and (best is None or competitor.price < best.price):
            best, best_query = competitor, query
        if best and index >= 1:
            break
        if index < len(queries) - 1:
            await pause()
    return best, best_query or (queries[-1] if queries else "")


async def search_competitors(page: Page, products: list[ClientProduct]) -> list[Result]:
    console.rule("[bold cyan]Крок 2: пошук мінімальних цін конкурентів на Prom")
    results: list[Result] = []
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Товари", total=len(products))
        for index, product in enumerate(products, start=1):
            progress.update(task, description=f"Товар {index}/{len(products)}")
            result = Result(product=product)
            try:
                result.competitor, result.query = await find_competitor(page, product)
                if result.competitor is None:
                    result.note = "Конкурентів не знайдено"
                    progress.console.print(
                        f"[yellow]--[/yellow] {product.name[:60]}: конкурентів не знайдено"
                    )
                else:
                    diff = product.price - result.competitor.price
                    color = "red" if diff > 0 else "green" if diff < 0 else "white"
                    progress.console.print(
                        f"[green]OK[/green] {product.name[:60]} | модель: {product.model or '-'} | "
                        f"клієнт {product.price:,.0f} | конкурент {result.competitor.price:,.0f} "
                        f"([{color}]{diff:+,.0f}[/{color}]) {result.competitor.shop}"
                    )
                    progress.console.print(f"   [dim]збіг: {result.competitor.title[:90]}[/dim]")
            except Exception as exc:  # noqa: BLE001
                result.note = f"Помилка: {type(exc).__name__}"
                progress.console.print(f"[red]ERR[/red] {product.name[:60]}: {type(exc).__name__}: {exc}")
            results.append(result)
            progress.advance(task)
            if index < len(products):
                await pause()
    return results


# ---------------------------------------------------------------------------
# Крок 3: Excel
# ---------------------------------------------------------------------------
def style_link(cell, url: str) -> None:
    cell.value = url
    cell.hyperlink = url
    cell.font = Font(color="0563C1", underline="single")


def write_report(results: list[Result], path: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Порівняння цін"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"

    for result in results:
        product, comp = result.product, result.competitor
        row_index = ws.max_row + 1
        if comp:
            diff = round(product.price - comp.price, 2)
            ws.append([product.name, product.price, comp.price, diff, comp.shop, comp.url, product.url])
        else:
            ws.append([product.name, product.price, None, None, result.note or "Конкурентів не знайдено", None, product.url])

        for cell in ws[row_index]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
        for col in (2, 3):
            ws.cell(row_index, col).number_format = "#,##0.00"
        diff_cell = ws.cell(row_index, 4)
        diff_cell.number_format = "+#,##0.00;-#,##0.00;0.00"
        diff_cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        if comp:
            if diff_cell.value > 0:
                diff_cell.fill = FILL_WORSE
            elif diff_cell.value < 0:
                diff_cell.fill = FILL_BETTER
            style_link(ws.cell(row_index, 6), comp.url)
        style_link(ws.cell(row_index, 7), product.url)

    for idx in range(1, len(HEADERS) + 1):
        limit = COL_WIDTH_LIMITS.get(idx, DEFAULT_COL_WIDTH_LIMIT)
        longest = 0
        for (value,) in ws.iter_rows(min_col=idx, max_col=idx, values_only=True):
            if value is None:
                continue
            text = f"{value:,.2f}" if isinstance(value, float) else str(value)
            longest = max(longest, len(text))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 2, 12), limit)

    target = Path(path)
    try:
        wb.save(target)
    except PermissionError:
        target = target.with_name(f"{target.stem}_backup{target.suffix}")
        console.print(f"[yellow]Файл {path} відкритий в іншій програмі. Зберігаю в {target}[/yellow]")
        wb.save(target)
    return target


def print_summary(results: list[Result], saved_to: Path) -> None:
    compared = [r for r in results if r.competitor]
    pricier = sum(1 for r in compared if r.product.price > r.competitor.price)
    cheaper = sum(1 for r in compared if r.product.price < r.competitor.price)
    table = Table(title="Підсумок")
    table.add_column("Показник")
    table.add_column("Значення", justify="right")
    table.add_row("Товарів клієнта", str(len(results)))
    table.add_row("Знайдено конкурента", str(len(compared)))
    table.add_row("Клієнт дорожчий", str(pricier))
    table.add_row("Клієнт дешевший", str(cheaper))
    table.add_row("Без конкурента", str(len(results) - len(compared)))
    table.add_row("Файл", str(saved_to))
    console.print(table)


# ---------------------------------------------------------------------------
# Точка входу
# ---------------------------------------------------------------------------
async def run() -> None:
    results: list[Result] = []
    async with async_playwright() as playwright:
        browser, context = await new_context(playwright)
        page = await context.new_page()
        try:
            products = await collect_client_products(page)
            if not products:
                console.print("[red]Товарів клієнта не знайдено. Перевірте TARGET_SHOP_URL / CATALOG_URL.[/red]")
                return
            results = await search_competitors(page, products)
        finally:
            if results:
                console.rule("[bold cyan]Крок 3: збереження звіту")
                print_summary(results, write_report(results, OUTPUT_FILE))
            try:
                await context.close()
                await browser.close()
            except Exception:  # noqa: BLE001
                pass


def main() -> None:
    try:
        asyncio.run(run())
    except KeyboardInterrupt:
        console.print(f"\n[yellow]Зупинено користувачем (Ctrl+C). Зібрана частина даних збережена в {OUTPUT_FILE}[/yellow]")
        sys.exit(130)


if __name__ == "__main__":
    main()
