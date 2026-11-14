"""Асинхронний парсер контактів продавців Prom.ua.

Етап 1: зі сторінок категорії збираються ТІЛЬКИ посилання на картки товарів.
Етап 2: кожна картка відкривається, з блоку продавця беруться назва компанії,
        сайт продавця (*.prom.ua), назва товару (h1) і телефони. За потреби
        додатково відкривається /ua/contacts сайту продавця (месенджери, email).
Етап 3: інкрементальне збереження в Excel з готовим текстом для розсилки.
"""
from __future__ import annotations

import asyncio
import json
import random
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
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

import config

console = Console()

HEADERS = [
    "Название магазина",
    "URL магазина",
    "Телефоны",
    "Telegram / Viber",
    "Email",
    "Пример товара",
    "Готовое сообщение для рассылки",
]

PRODUCT_URL_RE = re.compile(
    r"^https?://(?:www\.)?prom\.ua/(?:ua/|ru/)?p\d+(?:-[^/?#]*)?(?:[?#].*)?$", re.I
)
COMPANY_HREF_RE = re.compile(r"/c(\d+)-([a-z0-9\-]*)\.html", re.I)
PHONE_RE = re.compile(
    r"(?<!\d)(?:\+?\s*38)?[\s\-\.]*\(?0\d{2}\)?[\s\-\.]*\d{3}[\s\-\.]*\d{2}[\s\-\.]*\d{2}(?!\d)"
)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")
MESSENGER_RE = re.compile(
    r"https?://(?:www\.)?(?:t\.me|telegram\.me|viber\.click|msng\.link|wa\.me)/[^\s\"'<>\\)\]]+",
    re.I,
)
BAD_EMAIL_SUFFIXES = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js")
NOT_FOUND_TITLE_RE = re.compile(r"\b404\b|не знайден|не найден|not found", re.I)


# ---------------------------------------------------------------------------
# Моделі та утиліти
# ---------------------------------------------------------------------------
@dataclass
class Shop:
    """Продавець, знайдений на сторінці товару."""

    company_id: str
    name: str
    url: str                 # що йде в колонку «URL магазина»
    site_url: str = ""       # сайт продавця (*.prom.ua), якщо є
    product_url: str = ""
    product_name: str = ""
    phones: list[str] = field(default_factory=list)

    @property
    def label(self) -> str:
        return self.name or self.company_id


@dataclass
class Contacts:
    name: str = ""
    phones: list[str] = field(default_factory=list)
    messengers: list[str] = field(default_factory=list)
    emails: list[str] = field(default_factory=list)


async def pause(min_s: float | None = None, max_s: float | None = None) -> None:
    """Випадкова пауза між переходами."""
    await asyncio.sleep(random.uniform(min_s or config.MIN_DELAY, max_s or config.MAX_DELAY))


def ordered_unique(items) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


def clean_text(value: str | None, limit: int = 200) -> str:
    return re.sub(r"\s+", " ", value or "").strip()[:limit]


def normalize_phone(raw: str) -> str | None:
    """Приводить український номер до формату +380XXXXXXXXX."""
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 12 and digits.startswith("380"):
        national = digits[3:]
    elif len(digits) == 11 and digits.startswith("80"):
        national = digits[2:]
    elif len(digits) == 10 and digits.startswith("0"):
        national = digits[1:]
    else:
        return None
    # 0-800 - безкоштовні лінії (зокрема підтримка Prom), не потрібні
    if len(national) != 9 or national[0] == "0" or national.startswith("800"):
        return None
    return "+380" + national


def extract_phones(texts: list[str]) -> list[str]:
    phones: list[str] = []
    for text in texts:
        phone = normalize_phone(text)
        if phone:
            phones.append(phone)
        for match in PHONE_RE.finditer(text):
            phone = normalize_phone(match.group(0))
            if phone:
                phones.append(phone)
    return ordered_unique(phones)


def extract_emails(texts: list[str]) -> list[str]:
    emails: list[str] = []
    for text in texts:
        for match in EMAIL_RE.findall(text):
            email = match.lower().strip(".")
            if email.endswith(BAD_EMAIL_SUFFIXES) or email.endswith("prom.ua"):
                continue
            emails.append(email)
    return ordered_unique(emails)


def is_prom_messenger(link: str) -> bool:
    """Чи це месенджер/бот самого Prom (його треба виключити)."""
    lowered = link.lower()
    return any(bad in lowered for bad in config.PROM_MESSENGER_BLACKLIST)


def extract_messengers(texts: list[str]) -> list[str]:
    links: list[str] = []
    for text in texts:
        for match in MESSENGER_RE.findall(text):
            link = match.rstrip(".,;")
            if "/share" in link.lower() or is_prom_messenger(link):
                continue
            links.append(link)
    return ordered_unique(links)


def _host(url: str) -> str:
    try:
        return (urlsplit(url).hostname or "").lower()
    except ValueError:
        return ""


def is_service_prom_host(host: str) -> bool:
    """True для службових піддоменів Prom (sprava, premiya, gtw, my, support ...)."""
    if not host.endswith(".prom.ua"):
        return False
    sub = host[: -len(".prom.ua")]
    return any(part in config.SERVICE_SUBDOMAINS for part in sub.split("."))


def shop_subdomain(url: str) -> str | None:
    """Повертає піддомен магазину з URL або None, якщо це не магазин."""
    host = _host(url)
    if not host.endswith(".prom.ua"):
        return None
    sub = host[: -len(".prom.ua")]
    if not sub or "." in sub or sub in config.SERVICE_SUBDOMAINS:
        return None
    return sub


def site_base_url(url: str) -> str:
    """Нормалізує сайт продавця до https://host. Порожній рядок, якщо це не сайт магазину."""
    if not url:
        return ""
    host = _host(url)
    if not host or host in {"prom.ua", "www.prom.ua"} or is_service_prom_host(host):
        return ""
    if host.endswith(".prom.ua") and shop_subdomain(url) is None:
        return ""
    return f"https://{host}"


def page_url(base: str, page_number: int) -> str:
    parts = urlsplit(base)
    query = dict(parse_qsl(parts.query))
    query["page"] = str(page_number)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), ""))


# Картка товару (/p123) або сторінка компанії (/c123) на самому prom.ua, не домен магазину.
MARKETPLACE_ITEM_RE = re.compile(r"/(?:ua|ru)/[pc]\d+", re.I)

# Мобільні коди України. Решта +380.. вважається стаціонарним.
MOBILE_PREFIXES = (
    "+38050", "+38063", "+38066", "+38067", "+38068", "+38073",
    "+38093", "+38095", "+38096", "+38097", "+38098", "+38099",
)


def is_marketplace_item_url(url: str) -> bool:
    """True, якщо в колонку URL потрапила картка товару чи сторінка компанії на prom.ua."""
    host = _host(url)
    if host not in {"prom.ua", "www.prom.ua"}:
        return False
    try:
        path = urlsplit(url).path
    except ValueError:
        return True
    return bool(MARKETPLACE_ITEM_RE.search(path))


def is_shop_domain(url: str) -> bool:
    """Домен самого магазину: власний сайт або піддомен *.prom.ua, але не prom.ua/p|c."""
    host = _host(url)
    if not host or host in {"prom.ua", "www.prom.ua"} or is_service_prom_host(host):
        return False
    return not is_marketplace_item_url(url)


def is_mobile_phone(phone: str) -> bool:
    return phone.startswith(MOBILE_PREFIXES)


def split_cell(value) -> list[str]:
    if not value:
        return []
    return [line.strip() for line in str(value).split("\n") if line.strip()]


def lead_tier(phones: list[str], messengers: list[str]) -> int:
    """0 - месенджер і мобільний, 1 - мобільний, 2 - месенджер без мобільного, 3 - лише стаціонарний."""
    has_mobile = any(is_mobile_phone(phone) for phone in phones)
    if messengers and has_mobile:
        return 0
    if has_mobile:
        return 1
    if messengers:
        return 2
    return 3


def is_keepable_lead(url: str, phones: list[str], messengers: list[str]) -> bool:
    """Лід лишається, лише якщо є домен магазину і хоча б телефон або Telegram/Viber."""
    return is_shop_domain(url) and bool(phones or messengers)


def build_message(shop_name: str, sample_product: str) -> str:
    return config.MESSAGE_TEMPLATE.format(
        shop_name=shop_name,
        sample_product=sample_product or config.DEFAULT_SAMPLE_PRODUCT,
    )


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
class LeadsWriter:
    """Інкрементальний запис лідів у xlsx після кожного магазину."""

    MAX_WIDTHS = {7: 100}
    DEFAULT_MAX_WIDTH = 60

    def __init__(self, path: str) -> None:
        self.path = Path(path)
        self.known_urls: set[str] = set()
        if self.path.exists():
            self.wb = load_workbook(self.path)
            self.ws = self.wb.active
            self._purge_bad_rows()
            for row in self.ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
                if row[0]:
                    self.known_urls.add(str(row[0]).strip())
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
            self.ws.title = "Leads"
            self.ws.append(HEADERS)
            self._style_header()
            self._save()

    def _row_kept(self, row) -> bool:
        url = str(row[1] or "")
        return is_keepable_lead(url, split_cell(row[2]), split_cell(row[3]))

    def _purge_bad_rows(self) -> None:
        """Прибирає службові піддомени, посилання prom.ua/p|c і рядки без телефону та месенджера."""
        to_delete = [
            idx
            for idx, row in enumerate(
                self.ws.iter_rows(min_row=2, max_col=len(HEADERS), values_only=True), start=2
            )
            if any(row) and not self._row_kept(row)
        ]
        for idx in reversed(to_delete):
            self.ws.delete_rows(idx)
        if to_delete:
            console.print(f"[yellow]Прибрано непридатних рядків зі старого файлу: {len(to_delete)}[/yellow]")
            self._save()

    def _style_header(self) -> None:
        for cell in self.ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        self.ws.freeze_panes = "A2"

    def _autowidth(self) -> None:
        for idx in range(1, len(HEADERS) + 1):
            limit = self.MAX_WIDTHS.get(idx, self.DEFAULT_MAX_WIDTH)
            longest = 0
            for (value,) in self.ws.iter_rows(min_col=idx, max_col=idx, values_only=True):
                if value is None:
                    continue
                longest = max(longest, *(len(line) for line in str(value).split("\n")))
            self.ws.column_dimensions[get_column_letter(idx)].width = min(longest + 2, limit)

    def _save(self) -> None:
        self._autowidth()
        try:
            self.wb.save(self.path)
        except PermissionError:
            backup = self.path.with_name(f"{self.path.stem}_backup{self.path.suffix}")
            console.print(
                f"[yellow]Файл {self.path} відкритий в іншій програмі. "
                f"Зберігаю в {backup}[/yellow]"
            )
            self.wb.save(backup)

    def add(self, shop: Shop, contacts: Contacts) -> bool:
        if not is_keepable_lead(shop.url, contacts.phones, contacts.messengers):
            return False
        name = contacts.name or shop.name or shop.company_id
        row = [
            name,
            shop.url,
            "\n".join(contacts.phones),
            "\n".join(contacts.messengers),
            "\n".join(contacts.emails),
            shop.product_name,
            build_message(name, shop.product_name),
        ]
        self.ws.append(row)
        for cell in self.ws[self.ws.max_row]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        self.known_urls.add(shop.url)
        self._save()
        return True

    def sort_leads(self) -> int:
        """Залишає придатні ліди і ставить нагору ті, де є Telegram/Viber і мобільний."""
        kept = [
            row
            for row in self.ws.iter_rows(min_row=2, max_col=len(HEADERS), values_only=True)
            if any(row) and self._row_kept(row)
        ]
        kept.sort(key=lambda row: (
            lead_tier(split_cell(row[2]), split_cell(row[3])),
            -len(split_cell(row[2])),
            str(row[0] or "").lower(),
        ))
        if self.ws.max_row > 1:
            self.ws.delete_rows(2, self.ws.max_row - 1)
        for row in kept:
            self.ws.append(list(row))
            for cell in self.ws[self.ws.max_row]:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        self.known_urls = {str(row[1]).strip() for row in kept if row[1]}
        self._style_header()
        self._save()
        return len(kept)


# ---------------------------------------------------------------------------
# Браузер
# ---------------------------------------------------------------------------
async def block_heavy_resources(route: Route) -> None:
    if route.request.resource_type in config.BLOCKED_RESOURCE_TYPES:
        await route.abort()
    else:
        await route.continue_()


async def new_context(playwright) -> tuple[object, BrowserContext]:
    browser = await playwright.chromium.launch(
        headless=config.HEADLESS,
        args=["--disable-blink-features=AutomationControlled"],
    )
    context = await browser.new_context(
        user_agent=config.USER_AGENT,
        locale=config.LOCALE,
        viewport=config.VIEWPORT,
    )
    context.set_default_timeout(config.ACTION_TIMEOUT_MS)
    context.set_default_navigation_timeout(config.NAV_TIMEOUT_MS)
    await context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    await context.route("**/*", block_heavy_resources)
    return browser, context


async def smooth_scroll(page: Page) -> None:
    """Плавна прокрутка вниз для підвантаження карток товарів."""
    last_height = 0
    stagnant = 0
    for _ in range(config.SCROLL_MAX_STEPS):
        step = random.randint(config.SCROLL_STEP_MIN_PX, config.SCROLL_STEP_MAX_PX)
        await page.mouse.wheel(0, step)
        await asyncio.sleep(random.uniform(0.3, 0.8))
        metrics = await page.evaluate(
            "() => ({h: document.documentElement.scrollHeight,"
            " y: window.scrollY + window.innerHeight})"
        )
        at_bottom = metrics["y"] >= metrics["h"] - 20
        if at_bottom:
            stagnant = stagnant + 1 if metrics["h"] == last_height else 0
            if stagnant >= 2:
                break
        last_height = metrics["h"]


# ---------------------------------------------------------------------------
# Етап 1: збір посилань на картки товарів
# ---------------------------------------------------------------------------
LISTING_LINKS_JS = "() => [...document.querySelectorAll(\"a[href*='/p']\")].map(a => a.href)"


def clean_product_url(href: str) -> str | None:
    """Повертає URL картки товару без query/fragment або None."""
    if not PRODUCT_URL_RE.match(href):
        return None
    parts = urlsplit(href)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))


async def extract_product_links(page: Page, limit: int) -> list[str]:
    hrefs = await page.evaluate(LISTING_LINKS_JS)
    return ordered_unique(clean_product_url(h) for h in hrefs)[:limit]


async def collect_product_links(page: Page) -> list[str]:
    products: list[str] = []
    console.rule("[bold cyan]Етап 1: збір посилань на товари")
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Сторінки видачі", total=config.MAX_PAGES)
        for number in range(1, config.MAX_PAGES + 1):
            url = page_url(config.CATEGORY_URL, number)
            try:
                response = await page.goto(url, wait_until="domcontentloaded")
                if response and response.status >= 400:
                    progress.console.print(f"[yellow]Сторінка {number}: HTTP {response.status}, пропуск[/yellow]")
                    progress.advance(task)
                    continue
                await smooth_scroll(page)
                links = await extract_product_links(page, config.MAX_PRODUCTS_PER_PAGE)
                if not links:
                    progress.console.print(
                        f"[yellow]Сторінка {number}: товарів немає. Можлива капча, чекаю "
                        f"{config.CAPTCHA_WAIT_SEC} с...[/yellow]"
                    )
                    await asyncio.sleep(config.CAPTCHA_WAIT_SEC)
                    await smooth_scroll(page)
                    links = await extract_product_links(page, config.MAX_PRODUCTS_PER_PAGE)
                before = len(products)
                products = ordered_unique([*products, *links])
                progress.console.print(
                    f"Сторінка {number}: нових товарів [green]{len(products) - before}[/green], "
                    f"всього [bold]{len(products)}[/bold]"
                )
            except Exception as exc:  # noqa: BLE001
                progress.console.print(f"[red]Сторінка {number}: помилка {type(exc).__name__}: {exc}[/red]")
            progress.advance(task)
            if number < config.MAX_PAGES:
                await pause()
    return products


# ---------------------------------------------------------------------------
# Етап 2: картка товару -> продавець -> контакти
# ---------------------------------------------------------------------------
REVEAL_TEXT_RE = re.compile(r"^\s*Показати(\s+(номер|телефон)\w*)*\s*$", re.I)
REVEAL_ATTR_SELECTORS = [
    '[data-qaid*="show_phone"]',
    '[data-qaid*="phone"][role="button"]',
    '[data-qaid*="phone"] button',
    '[class*="show-phone"]',
    '[class*="showPhone"]',
    '[class*="ShowPhone"]',
    '[class*="phone"][role="button"]',
    "button[data-phone]",
]
MAX_REVEAL_CLICKS = 8

PRODUCT_DATA_JS = """
() => {
  const scope = document.querySelector('[data-qaid="company_dashboard"]') || document;
  const link = scope.querySelector('a[data-qaid="company_name"]')
            || document.querySelector('a[data-qaid="company_name"]');
  const names = [...document.querySelectorAll('[data-qaid="company_name"]')]
                  .map(e => (e.textContent || '').trim());
  const h1 = document.querySelector('h1');
  return {
    h1: h1 ? h1.textContent : '',
    companyHref: link ? link.href : '',
    companyText: link ? (link.textContent || '').trim() : '',
    companyNames: names
  };
}
"""

DOM_PHONES_JS = """
() => {
  const popup = document.querySelector('[data-qaid="contacts_popup"]');
  return {
    tel: [...document.querySelectorAll('a[href^="tel:"]')].map(a => a.getAttribute('href')),
    popup: popup ? popup.innerText : ''
  };
}
"""

SITE_DATA_JS = """
() => {
  const meta = document.querySelector('meta[property="og:site_name"]');
  return {
    tel: [...document.querySelectorAll('a[href^="tel:"]')].map(a => a.getAttribute('href')),
    hrefs: [...document.querySelectorAll('a[href]')].map(a => a.href),
    text: document.body ? document.body.innerText : '',
    ogName: meta ? meta.content : '',
    title: document.title || ''
  };
}
"""


async def reveal_phones(page: Page) -> int:
    """Натискає кнопки «Показати» / «Показати номер» тощо."""
    clicks = 0
    text_locators = [
        page.get_by_role("button", name=REVEAL_TEXT_RE),
        page.get_by_text(REVEAL_TEXT_RE),
    ]
    for locator in text_locators:
        for _ in range(MAX_REVEAL_CLICKS):
            try:
                if await locator.count() == 0:
                    break
                await locator.first.click(timeout=2_000)
                clicks += 1
                await asyncio.sleep(random.uniform(0.3, 0.7))
            except Exception:  # noqa: BLE001
                break
    for selector in REVEAL_ATTR_SELECTORS:
        try:
            handles = await page.query_selector_all(selector)
        except Exception:  # noqa: BLE001
            continue
        for handle in handles[:MAX_REVEAL_CLICKS]:
            try:
                await handle.click(timeout=1_500)
                clicks += 1
                await asyncio.sleep(random.uniform(0.2, 0.5))
            except Exception:  # noqa: BLE001
                continue
    if clicks:
        await asyncio.sleep(random.uniform(0.8, 1.5))
    return clicks


def _json_str(value: str) -> str:
    try:
        return json.loads(f'"{value}"')
    except ValueError:
        return value


def parse_company_state(html: str, company_id: str) -> tuple[str, str, list[str]]:
    """Витягує з вбудованого JSON-стану сторінки: назву, сайт (webSiteUrl) і телефони компанії."""
    html = html.replace("\\u002F", "/").replace("\\/", "/")
    name = site = ""
    anchor = html.find('"company":{"id":' + company_id + ",")
    if anchor >= 0:
        segment = html[anchor : anchor + 4000]
        if m := re.search(r'"webSiteUrl":"([^"]*)"', segment):
            site = m.group(1)
        if m := re.search(r'"name":"([^"]*)"', segment):
            name = _json_str(m.group(1))

    phones_block = ""
    m = re.search(
        r'"combinedOpinionStats":\{"id":' + company_id + r',[^}]*\},"phones":\[(.*?)\]',
        html,
        re.S,
    )
    if m:
        phones_block = m.group(1)
    elif anchor >= 0:
        m = re.search(r'"phones":\[(.*?)\]', html[anchor : anchor + 20000], re.S)
        phones_block = m.group(1) if m else ""
    phones = extract_phones(re.findall(r'"number":"([^"]+)"', phones_block))
    return name, site, phones


async def open_contacts_popup(page: Page) -> None:
    """Натискає «Контакти» в блоці продавця, потім кнопки «Показати номер»."""
    try:
        await page.locator('[data-qaid="contacts_btn"]').first.click(timeout=3_000)
        await asyncio.sleep(random.uniform(0.6, 1.2))
    except Exception:  # noqa: BLE001
        pass
    await reveal_phones(page)


async def parse_product_page(page: Page, product_url: str) -> Shop | None:
    """Відкриває картку товару й повертає дані продавця. None, якщо продавця не знайдено."""
    await page.goto(product_url, wait_until="domcontentloaded")
    try:
        await page.wait_for_selector('[data-qaid="company_name"]', timeout=8_000)
    except Exception:  # noqa: BLE001
        pass

    data = await page.evaluate(PRODUCT_DATA_JS)
    html = await page.content()

    match = COMPANY_HREF_RE.search(data["companyHref"])
    if not match:
        match = re.search(r'"company":\{"id":(\d+),', html)
        if not match:
            return None
        company_id, slug = match.group(1), ""
    else:
        company_id, slug = match.group(1), match.group(2)

    state_name, state_site, state_phones = parse_company_state(html, company_id)

    name = ""
    for candidate in data["companyNames"] or [data["companyText"]]:
        candidate = re.sub(r"^Продавець\s+", "", candidate or "").strip()
        if candidate:
            name = candidate
            break
    name = name or state_name

    # Кнопка «Контакти» / «Показати номер» + a[href^="tel:"] у DOM
    await open_contacts_popup(page)
    dom = await page.evaluate(DOM_PHONES_JS)
    phones = ordered_unique([*state_phones, *extract_phones([*dom["tel"], dom["popup"]])])

    site_url = site_base_url(state_site)
    company_page = f"https://prom.ua/ua/c{company_id}-{slug}.html" if slug else f"https://prom.ua/ua/c{company_id}.html"
    return Shop(
        company_id=company_id,
        name=clean_text(name, 120),
        url=site_url or company_page,
        site_url=site_url,
        product_url=product_url,
        product_name=clean_text(data["h1"]),
        phones=phones,
    )


def guess_shop_name(og_name: str, title: str) -> str:
    if clean_text(og_name) and not re.search(r"prom\.ua", og_name, re.I):
        return clean_text(og_name, 120)
    for part in re.split(r"\s[-|·]\s", title or ""):
        part = clean_text(part, 120)
        if part and not re.match(r"контакт", part, re.I) and not re.search(r"prom\.ua", part, re.I):
            return part
    return ""


async def parse_site_page(page: Page) -> Contacts:
    await reveal_phones(page)
    data = await page.evaluate(SITE_DATA_JS)
    html = await page.content()
    mailtos = [h[len("mailto:"):] for h in data["hrefs"] if h.lower().startswith("mailto:")]
    return Contacts(
        name=guess_shop_name(data["ogName"], data["title"]),
        phones=extract_phones([*data["tel"], data["text"]]),
        messengers=extract_messengers([*data["hrefs"], html]),
        emails=extract_emails([*mailtos, data["text"]]),
    )


async def scrape_site_contacts(page: Page, site_url: str) -> Contacts:
    """Сторінка контактів сайту продавця: /ua/contacts -> /contacts -> головна."""
    candidates = [f"{site_url}/ua/contacts", f"{site_url}/contacts", site_url]
    result = Contacts()
    for url in candidates:
        response = await page.goto(url, wait_until="domcontentloaded")
        status = response.status if response else 0
        if status == 404 or status >= 500 or NOT_FOUND_TITLE_RE.search(await page.title()):
            await pause(0.8, 1.8)
            continue
        await asyncio.sleep(random.uniform(0.8, 1.5))
        found = await parse_site_page(page)
        result.name = result.name or found.name
        result.phones = ordered_unique([*result.phones, *found.phones])
        result.messengers = ordered_unique([*result.messengers, *found.messengers])
        result.emails = ordered_unique([*result.emails, *found.emails])
        if result.phones:
            break
        await pause(0.8, 1.8)
    return result


async def collect_leads(page: Page, product_urls: list[str], writer: LeadsWriter) -> dict[str, int]:
    console.rule("[bold cyan]Етап 2-3: продавці, контакти та збереження в Excel")
    stats = {
        "saved": 0,
        "with_phones": 0,
        "failed": 0,
        "duplicates": 0,
        "skipped_url": 0,
        "skipped_empty": 0,
    }
    seen_companies: set[str] = set()

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Товари", total=len(product_urls))
        for index, product_url in enumerate(product_urls):
            if config.MAX_SHOPS and stats["saved"] >= config.MAX_SHOPS:
                progress.console.print("[dim]Досягнуто MAX_SHOPS, зупиняюсь[/dim]")
                break
            delay = (config.MIN_DELAY, config.MAX_DELAY)
            progress.update(task, description=f"Товар {index + 1}/{len(product_urls)}")
            try:
                shop = await parse_product_page(page, product_url)
                if shop is None:
                    stats["failed"] += 1
                    progress.console.print(f"[yellow]SKIP[/yellow] {product_url}: продавця не знайдено")
                elif shop.company_id in seen_companies or shop.url in writer.known_urls:
                    stats["duplicates"] += 1
                    delay = (1.0, 2.0)
                else:
                    seen_companies.add(shop.company_id)
                    contacts = Contacts(name=shop.name, phones=list(shop.phones))
                    if shop.site_url:
                        await pause(1.0, 2.0)
                        try:
                            site = await scrape_site_contacts(page, shop.site_url)
                            contacts.name = contacts.name or site.name
                            contacts.phones = ordered_unique([*contacts.phones, *site.phones])
                            contacts.messengers = site.messengers
                            contacts.emails = site.emails
                        except Exception as exc:  # noqa: BLE001
                            progress.console.print(
                                f"[yellow]WARN[/yellow] сайт {shop.site_url}: {type(exc).__name__}, "
                                f"лишаю дані з картки товару"
                            )
                    if not is_shop_domain(shop.url):
                        stats["skipped_url"] += 1
                        progress.console.print(
                            f"[dim]SKIP[/dim] {shop.label}: URL не домен магазину ({shop.url})"
                        )
                    elif not contacts.phones and not contacts.messengers:
                        stats["skipped_empty"] += 1
                        progress.console.print(
                            f"[dim]SKIP[/dim] {shop.label}: немає телефону і Telegram/Viber"
                        )
                    elif writer.add(shop, contacts):
                        stats["saved"] += 1
                        stats["with_phones"] += bool(contacts.phones)
                        progress.console.print(
                            f"[green]OK[/green] {shop.label}: тел. {len(contacts.phones)}, "
                            f"месенджерів {len(contacts.messengers)}, email {len(contacts.emails)}"
                        )
            except Exception as exc:  # noqa: BLE001
                stats["failed"] += 1
                progress.console.print(f"[red]ERR[/red] {product_url}: {type(exc).__name__}: {exc}")
            progress.advance(task)
            if index < len(product_urls) - 1:
                await pause(*delay)
    return stats


# ---------------------------------------------------------------------------
# Точка входу
# ---------------------------------------------------------------------------
async def run() -> None:
    writer = LeadsWriter(config.OUTPUT_FILE)
    async with async_playwright() as playwright:
        browser, context = await new_context(playwright)
        page = await context.new_page()
        try:
            products = await collect_product_links(page)
            if not products:
                console.print("[red]Товарів не знайдено. Перевірте CATEGORY_URL або пройдіть капчу.[/red]")
                return
            stats = await collect_leads(page, products, writer)
            kept = writer.sort_leads()

            table = Table(title="Підсумок")
            table.add_column("Показник")
            table.add_column("Значення", justify="right")
            table.add_row("Оброблено товарів", str(len(products)))
            table.add_row("Збережено продавців", str(stats["saved"]))
            table.add_row("З телефонами", str(stats["with_phones"]))
            table.add_row("Пропущено без домену", str(stats["skipped_url"]))
            table.add_row("Пропущено без контакту", str(stats["skipped_empty"]))
            table.add_row("Дублікати продавців", str(stats["duplicates"]))
            table.add_row("Помилок/пропусків", str(stats["failed"]))
            table.add_row("У файлі після сортування", str(kept))
            table.add_row("Файл", config.OUTPUT_FILE)
            console.print(table)
        finally:
            try:
                writer.sort_leads()
            except Exception:  # noqa: BLE001
                pass
            try:
                await context.close()
                await browser.close()
            except Exception:  # noqa: BLE001
                pass


def main() -> None:
    try:
        asyncio.run(run())
    except KeyboardInterrupt:
        console.print(
            f"\n[yellow]Зупинено користувачем (Ctrl+C). "
            f"Зібрані дані збережено в {config.OUTPUT_FILE}[/yellow]"
        )
        sys.exit(130)


if __name__ == "__main__":
    main()
