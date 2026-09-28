"""
Website file downloader using Playwright.
Files are saved with the real document name shown on the website.

Works in two ways:
  1. Script-based sites (like eci.gov.in): finds every document card, downloads it,
     and goes through all pages.
  2. Normal sites: finds file links (.pdf, .xlsx ...) in the page.

Setup (one time):
    pip install playwright pypdf
    playwright install chromium

Run:
    python download_all.py
"""

import io
import re
import mimetypes
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

import urllib.request


def detect_system_proxy():
    """Read the proxy Windows/the OS is already using (if any)."""
    proxies = urllib.request.getproxies()
    return proxies.get("https") or proxies.get("http")


SYSTEM_PROXY = detect_system_proxy()

# Paste your website links here
URLS = ["https://elibrary.sansad.in/search?query=Parliamentary%20Committee%20Reports&scope=571cd23f-4973-410f-a639-dabb4cbd805b&spc.sf=keyword_count&spc.sd=DESC"
]

# Where the name comes from in normal-link mode (tried in this order):
#   "row_text"  = title text in the same table row / list item as the link
#   "link_text" = the text of the link itself
#   "pdf_title" = the title stored inside the PDF file
NAME_ORDER = ["row_text", "link_text", "pdf_title"]

DOWNLOAD_DIR = Path("pdf_downloads")
HEADLESS = False  # False = browser window is visible, True = runs in background
OVERWRITE = True  # True = replace existing file, False = skip if it already exists
FILE_TYPES = ["pdf", "xlsx", "xls", "csv", "doc", "docx", "ppt", "pptx", "zip", "json"]
TIMEOUT = 30000  # milliseconds
CLICK_WAIT_SECONDS = 10  # how long to wait after clicking a card
MAX_PAGES = 100  # safety limit for the number of pages to go through
SAVE_DEBUG_HTML = True  # saves debug_page.html (the page as the browser sees it)
MAX_NAME_LENGTH = 150

GENERIC_NAMES = {"title", "untitled", "document", "download", "pdf", "file", "view", "click here"}

EXT_PATTERN = re.compile(r"\.(" + "|".join(FILE_TYPES) + r")(\?.*)?$", re.I)

# ---------------------------------------------------------------------------
# Browser-side scripts
# ---------------------------------------------------------------------------

# Normal-link mode: every link plus the text of its row
LINK_SCRIPT = """
els => els.map(e => {
    const linkText = (e.innerText || e.getAttribute('title') || e.getAttribute('aria-label') || '').trim();
    const row = e.closest('tr') || e.closest('li') || e.closest('[role=row]');
    let rowText = '';
    if (row) {
        const cells = Array.from(row.querySelectorAll('td, th'));
        const parts = cells.length ? cells : Array.from(row.children);
        const texts = parts
            .filter(c => !c.contains(e))
            .map(c => (c.innerText || '').replace(/\\s+/g, ' ').trim())
            .filter(t => t && !/^\\d+\\.?$/.test(t) && !/^\\(?[\\d.,]+\\s*(KB|MB|GB|B)\\)?$/i.test(t));
        texts.sort((a, b) => b.length - a.length);
        rowText = texts[0] || '';
    }
    return { href: e.href, linkText: linkText, rowText: rowText };
})
"""

# Card mode: find every card that shows "PDF | size", tag it,
# and return its title, its file link(s), and how many files it claims to have
CARD_SCRIPT = """
() => {
  document.querySelectorAll('[data-card-idx]').forEach(e => e.removeAttribute('data-card-idx'));
  const metaRe = /(PDF|DOCX?|XLSX?|ZIP)\\s*\\|/i;
  const hasMeta = el => metaRe.test(el.innerText || '');
  const isMetaLine = l => /^Date\\s*:/i.test(l) || /^(PDF|DOCX?|XLSX?|ZIP)\\s*\\|/i.test(l);
  const all = Array.from(document.body.querySelectorAll('*'));
  const metas = all.filter(el => hasMeta(el) && !Array.from(el.children).some(c => hasMeta(c)));

  // Some sites split one file-info row into several small sibling chips, e.g.
  // "PDF | 17 KB" and "PDF | 1 MB" side by side for a card that holds two files.
  // Climb each chip up while its container is still just a short chips row (not a
  // whole new card with its own title), and group same-row chips together so they
  // don't get mistaken for two separate (blank-title) cards.
  const groupKey = el => {
    let node = el;
    for (let i = 0; i < 3 && node.parentElement; i++) {
      const parent = node.parentElement;
      const text = (parent.innerText || '').trim();
      if (text.length > 200) break;
      node = parent;
    }
    return node;
  };
  const groupMap = new Map();
  metas.forEach(m => {
    const key = groupKey(m);
    if (!groupMap.has(key)) groupMap.set(key, []);
    groupMap.get(key).push(m);
  });
  const metaGroups = Array.from(groupMap.entries()).map(([key, ms]) => ms.length > 1 ? key : ms[0]);

  const cards = [];
  metaGroups.forEach(meta => {
    let card = meta;
    while (card.parentElement && card.parentElement !== document.body) {
      const parent = card.parentElement;
      if (metaGroups.filter(m => parent.contains(m)).length > 1) break;
      card = parent;
    }
    if (!cards.includes(card)) cards.push(card);
  });

  return cards.map((card, i) => {
    card.setAttribute('data-card-idx', String(i));
    const lines = (card.innerText || '').split('\\n').map(s => s.trim()).filter(Boolean);
    const metaAt = lines.findIndex(isMetaLine);
    let title = metaAt > 0 ? lines.slice(0, metaAt).join(' ') : '';
    if (!title) {
      const rest = lines.filter(l => !isMetaLine(l));
      rest.sort((a, b) => b.length - a.length);
      title = rest[0] || '';
    }
    const as = Array.from(card.querySelectorAll('a[href]'))
      .filter(a => a.href && !a.href.startsWith('javascript') && !a.href.endsWith('#'));
    const hrefs = [...new Set(as.map(a => a.href))];
    const fileMatches = (card.innerText || '').match(/(PDF|DOCX?|XLSX?|ZIP)\\s*\\|/gi) || [];
    const fileCount = Math.max(fileMatches.length, hrefs.length, 1);
    return { index: i, title: title, hrefs: hrefs, fileCount: fileCount };
  });
}
"""

# Pagination: find the enabled "next page" / "load more" control and tag it
NEXT_SCRIPT = """
() => {
  document.querySelectorAll('[data-next-page]').forEach(e => e.removeAttribute('data-next-page'));
  const visible = el => el.getClientRects().length > 0;
  const disabled = el => {
    for (let n = el, k = 0; n && k < 3; n = n.parentElement, k++) {
      const cls = (n.className && n.className.toString) ? n.className.toString() : '';
      if (n.disabled || n.getAttribute('aria-disabled') === 'true' || /disabled/i.test(cls)) return true;
    }
    return false;
  };
  const nearPageNumber = el => {
    let p = el;
    for (let k = 0; k < 3 && p.parentElement; k++) {
      p = p.parentElement;
      if ((p.innerText || p.textContent || '').length > 60) return false;
      const hit = Array.from(p.querySelectorAll('*')).some(
        x => x.children.length === 0 && /^\\d+$/.test((x.innerText || x.textContent || '').trim()));
      if (hit) return true;
    }
    return false;
  };
  const isNext = el => {
    const text = (el.innerText || el.textContent || '').trim();
    const label = ((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).toLowerCase();
    if (/^(next|next page|load more|show more|view more)$/i.test(text) || label.includes('next')) return true;
    return ['>', '\\u203a', '\\u00bb', '\\u2192'].includes(text) && nearPageNumber(el);
  };
  const els = Array.from(document.querySelectorAll('a, button, li, span, div, [role=button]'));
  const cands = els.filter(el => isNext(el) && visible(el) && !disabled(el));
  const deepest = cands.filter(el => !cands.some(o => o !== el && el.contains(o)));
  if (!deepest.length) return false;
  deepest[0].setAttribute('data-next-page', '1');
  return true;
}
"""

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def clean_filename(name):
    """Remove characters Windows does not allow and tidy the spaces."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    name = re.sub(r"\s+", " ", name).strip().rstrip(".")
    return name[:MAX_NAME_LENGTH].strip()


def is_generic(name):
    lowered = name.lower().strip()
    return (not lowered) or lowered in GENERIC_NAMES or lowered.startswith("microsoft word -")


def unique_in_run(folder, name, used):
    """If two different files get the same name in one run, add (2), (3)..."""
    path = folder / name
    n = 2
    while path.name.lower() in used:
        path = folder / f"{Path(name).stem} ({n}){Path(name).suffix}"
        n += 1
    used.add(path.name.lower())
    return path


def extension_for(url, content_type):
    match = EXT_PATTERN.search(url)
    if match:
        return "." + match.group(1).lower()
    return mimetypes.guess_extension((content_type or "").split(";")[0].strip()) or ".pdf"


def wait_for_content(page):
    """Wait until the script-based page has drawn its content."""
    page.wait_for_load_state("networkidle")
    try:
        page.wait_for_selector("text=/PDF/", timeout=15000)
    except PWTimeout:
        pass
    for _ in range(4):
        page.mouse.wheel(0, 2500)
        page.wait_for_timeout(400)
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(500)


# ---------------------------------------------------------------------------
# Normal links mode
# ---------------------------------------------------------------------------


def get_file_links(page):
    items = page.eval_on_selector_all("a[href]", LINK_SCRIPT)
    links, seen = [], set()
    for item in items:
        full = urljoin(page.url, item["href"])
        if EXT_PATTERN.search(full) and full not in seen:
            seen.add(full)
            links.append((full, item["rowText"], item["linkText"]))
    return links


# ---------------------------------------------------------------------------
# Item-page mode (search results / repository listings that link to a detail
# page for each record, e.g. DSpace-based digital libraries, instead of
# linking straight to a file)
# ---------------------------------------------------------------------------

ITEM_PAGE_HINTS = ("/items/", "/item/", "/handle/", "/record/", "/details/")


def looks_like_item_link(url):
    path = urlparse(url).path.lower()
    return any(hint in path for hint in ITEM_PAGE_HINTS)


def get_item_links(page):
    """Links on a search/results page that point to a record's own detail page."""
    items = page.eval_on_selector_all("a[href]", LINK_SCRIPT)
    links, seen = [], set()
    for item in items:
        full = urljoin(page.url, item["href"])
        if looks_like_item_link(full) and full not in seen:
            seen.add(full)
            links.append((full, item["rowText"], item["linkText"]))
    return links


# DSpace repositories (like elibrary.sansad.in) sometimes publish a
# citation_pdf_url meta tag that points at their INTERNAL server address
# (e.g. https://10.247.128.25:4000/...), which is unreachable from outside
# their own office network. The same file is also served on the site's own
# public domain at /server/api/core/bitstreams/<id>/content - rebuild that
# working link instead of trusting the broken one.
BITSTREAM_ID_RE = re.compile(r"/bitstreams/([0-9a-f-]{36})/", re.I)


def fix_internal_bitstream_url(url, page_url):
    match = BITSTREAM_ID_RE.search(url)
    if not match:
        return url
    origin = f"{urlparse(page_url).scheme}://{urlparse(page_url).netloc}"
    fixed = f"{origin}/server/api/core/bitstreams/{match.group(1)}/content"
    if fixed != url:
        print(f"  Fixed a broken internal file link -> {fixed}")
    return fixed


def find_download_link(page):
    """On a record's detail page, find the actual file: a citation_pdf_url meta tag
    (common on repository sites), a bitstream/download link, or a Download button."""
    meta_url = page.evaluate("""
        () => {
            const m = document.querySelector('meta[name="citation_pdf_url"]');
            return m ? m.content : '';
        }
    """)
    if meta_url:
        return fix_internal_bitstream_url(meta_url, page.url)

    ext_group = "|".join(FILE_TYPES)
    href = page.evaluate(
        """
        (extGroup) => {
            const extRe = new RegExp('\\\\.(' + extGroup + ')(\\\\?.*)?$', 'i');
            const links = Array.from(document.querySelectorAll('a[href]'));
            let a = links.find(l => /bitstream/i.test(l.href) && /download/i.test(l.href));
            if (!a) a = links.find(l => /^\\s*download\\s*$/i.test(l.innerText || ''));
            if (!a) a = links.find(l => extRe.test(l.href));
            return a ? a.href : '';
        }
        """,
        ext_group,
    )
    return fix_internal_bitstream_url(href, page.url) if href else None


def download_item_pages(context, page, item_links, folder, used):
    done, failed = 0, []
    for url, row_text, link_text in item_links:
        title = clean_filename(row_text or link_text) or "document"
        item_page = None
        try:
            item_page = context.new_page()
            item_page.set_default_timeout(TIMEOUT)
            item_page.goto(url, wait_until="domcontentloaded")
            file_url = find_download_link(item_page)
            item_page.close()
            item_page = None
            if not file_url:
                failed.append((title, "no file link found on item page"))
                continue

            response = context.request.get(urljoin(url, file_url), timeout=TIMEOUT)
            content_type = response.headers.get("content-type", "")
            if not response.ok:
                failed.append((title, f"HTTP {response.status}"))
                continue
            if "text/html" in content_type:
                failed.append((title, "opened a web page, not a file"))
                continue

            path = unique_in_run(folder, title + extension_for(file_url, content_type), used)
            path.write_bytes(response.body())
            print(f"  Downloaded: {path.name}")
            done += 1
        except Exception as e:
            failed.append((title, str(e).splitlines()[0] if str(e) else "error"))
        finally:
            if item_page:
                try:
                    item_page.close()
                except Exception:
                    pass

    print(f"  page done: downloaded {done}, not downloaded {len(failed)}")
    for title, reason in failed:
        print(f"   - {title[:80]}: {reason}")
    return done, failed


def go_to_next_search_page(page, before_urls):
    """Click the next-page control on a search/results listing (not card mode)."""
    if not page.evaluate(NEXT_SCRIPT):
        return False
    page.locator('[data-next-page="1"]').first.click(timeout=TIMEOUT)
    for _ in range(25):
        page.wait_for_timeout(300)
        current = [u for u, _, _ in get_item_links(page)]
        if current and current != before_urls:
            break
    page.wait_for_timeout(500)
    return True


def download_item_search_results(context, page, url, folder, used):
    page_no = 1
    item_links = get_item_links(page)
    while item_links and page_no <= MAX_PAGES:
        print(f"{url} -> page {page_no}: {len(item_links)} item pages found")
        download_item_pages(context, page, item_links, folder, used)
        before_urls = [u for u, _, _ in item_links]
        if not go_to_next_search_page(page, before_urls):
            break
        page_no += 1
        item_links = get_item_links(page)


def server_filename(response, url):
    header = response.headers.get("content-disposition", "")
    match = re.search(r"filename\*\s*=\s*[^']*''([^;]+)", header, re.I)
    if match:
        return unquote(match.group(1))
    match = re.search(r'filename\s*=\s*"?([^";]+)"?', header, re.I)
    if match:
        return unquote(match.group(1))
    return unquote(Path(urlparse(url).path).name)


def pdf_title(body):
    if PdfReader is None:
        return ""
    try:
        meta = PdfReader(io.BytesIO(body)).metadata
        return (meta.title or "").strip() if meta else ""
    except Exception:
        return ""


def choose_name(response, url, row_text, link_text):
    fallback = server_filename(response, url)
    ext = Path(fallback).suffix or extension_for(url, response.headers.get("content-type", ""))

    for source in NAME_ORDER:
        if source == "row_text":
            name = row_text
        elif source == "link_text":
            name = link_text
        else:
            name = pdf_title(response.body()) if ext.lower() == ".pdf" else ""

        name = clean_filename(name)
        if not is_generic(name):
            return name + ext

    return clean_filename(Path(fallback).stem) + ext


def download_normal_links(context, links, folder, used):
    for link, row_text, link_text in links:
        try:
            response = context.request.get(link, timeout=TIMEOUT)
            if not response.ok:
                print(f"  Failed (HTTP {response.status}): {link}")
                continue

            name = choose_name(response, link, row_text, link_text)
            path = unique_in_run(folder, name, used)
            if path.exists() and not OVERWRITE:
                print(f"  Skipped (already exists): {path.name}")
                continue

            path.write_bytes(response.body())
            print(f"  Downloaded: {path.name}")
        except Exception as e:
            print(f"  Failed: {link} ({e})")


# ---------------------------------------------------------------------------
# Cards mode (script-based sites)
# ---------------------------------------------------------------------------


def click_and_capture(context, page, target):
    """Click any locator and report what the site did: a download, a file response, a new tab, or a redirect."""
    downloads, popups, files = [], [], []

    def on_download(d):
        downloads.append(d)

    def on_popup(p):
        popups.append(p)
        p.on("download", on_download)

    def on_response(r):
        try:
            content_type = r.headers.get("content-type", "").lower()
        except Exception:
            return
        if any(k in content_type for k in ("pdf", "octet-stream", "msword", "officedocument", "ms-excel", "zip")):
            files.append(r.url)

    page.on("download", on_download)
    context.on("page", on_popup)
    context.on("response", on_response)
    start_url = page.url
    try:
        target.scroll_into_view_if_needed()
        target.click(timeout=TIMEOUT)

        for _ in range(CLICK_WAIT_SECONDS * 4):
            if downloads or files or popups or page.url != start_url:
                break
            page.wait_for_timeout(250)

        if popups and not downloads and not files:
            for _ in range(40):
                if downloads or files or popups[0].url not in ("", "about:blank"):
                    break
                page.wait_for_timeout(250)

        if downloads:
            result = ("download", downloads[0])
        elif files:
            result = ("url", files[0])
        elif popups:
            result = ("url", popups[0].url)
        elif page.url != start_url:
            result = ("url", page.url)
        else:
            result = (None, None)

        for popup in popups:
            try:
                popup.close()
            except Exception:
                pass
        if page.url != start_url:
            page.go_back()
            wait_for_content(page)
        return result
    finally:
        page.remove_listener("download", on_download)
        context.remove_listener("page", on_popup)
        context.remove_listener("response", on_response)


def click_card(context, page, index):
    """Click one card (its own link, or the card itself) and report what happened."""
    card = page.locator(f'[data-card-idx="{index}"]')
    target = card.locator("a").first if card.locator("a").count() else card
    return click_and_capture(context, page, target)


# Chip buttons like "PDF 1", "PDF 2" that a card shows for each of its files
CHIP_RE_TEMPLATE = r"^(PDF|DOC|DOCX|XLS|XLSX|ZIP|CSV|PPT|PPTX)\s*{n}$"


def download_numbered_chips(context, page, index, title, folder, used, already_ok, file_count):
    """Click 'PDF 1' / 'PDF 2' style chip buttons for files that have no direct link
    (used when a card's summary line advertises more files than we could grab directly)."""
    card = page.locator(f'[data-card-idx="{index}"]')
    downloaded = 0
    for n in range(already_ok + 1, file_count + 1):
        chip = card.get_by_text(re.compile(CHIP_RE_TEMPLATE.format(n=n), re.I), exact=True)
        if chip.count() == 0:
            continue
        try:
            kind, result = click_and_capture(context, page, chip.first)
        except Exception as e:
            print(f"  NOT downloaded: {title} (#{n}) ({e})")
            continue

        name = f"{title} ({n})" if file_count > 1 else title
        if kind == "download":
            ext = Path(result.suggested_filename).suffix or ".pdf"
            path = unique_in_run(folder, name + ext, used)
            result.save_as(path)
            print(f"  Downloaded: {path.name}")
            downloaded += 1
        elif kind == "url":
            try:
                response = context.request.get(result, timeout=TIMEOUT)
                content_type = response.headers.get("content-type", "")
                if response.ok and "text/html" not in content_type:
                    path = unique_in_run(folder, name + extension_for(result, content_type), used)
                    path.write_bytes(response.body())
                    print(f"  Downloaded: {path.name}")
                    downloaded += 1
            except Exception:
                pass
    return downloaded



def load_all_cards(page):
    """Scroll down until no more cards appear (for pages that load more while scrolling)."""
    last = -1
    for _ in range(30):
        count = len(page.evaluate(CARD_SCRIPT))
        if count == last:
            break
        last = count
        page.mouse.wheel(0, 4000)
        page.wait_for_timeout(700)
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(300)


def go_to_next_page(page):
    """Click the next-page control. Returns False when there is no next page."""
    before = [c["title"] for c in page.evaluate(CARD_SCRIPT)]
    if not page.evaluate(NEXT_SCRIPT):
        return False
    page.locator('[data-next-page="1"]').first.click(timeout=TIMEOUT)
    for _ in range(25):
        page.wait_for_timeout(300)
        titles = [c["title"] for c in page.evaluate(CARD_SCRIPT)]
        if titles and titles != before:
            break
    page.wait_for_timeout(500)
    return True


def process_card(context, page, index, card, folder, used):
    """Download every file inside one card. Returns (True, "") or (False, reason)."""
    title = clean_filename(card["title"]) or f"document {index + 1}"
    hrefs = card.get("hrefs", [])
    file_count = card.get("fileCount", 1)

    try:
        ok_count = 0
        last_reason = ""

        # 1) Direct link(s) inside the card (a folder/card can hold more than one file)
        for i, href in enumerate(hrefs):
            try:
                response = context.request.get(href, timeout=TIMEOUT)
                content_type = response.headers.get("content-type", "")
                if response.ok and "text/html" not in content_type:
                    name = title if file_count == 1 else f"{title} ({i + 1})"
                    path = unique_in_run(folder, name + extension_for(href, content_type), used)
                    path.write_bytes(response.body())
                    print(f"  Downloaded: {path.name}")
                    ok_count += 1
                else:
                    last_reason = f"HTTP {response.status} for {href}"
            except Exception as e:
                last_reason = str(e).splitlines()[0] if str(e) else "error"

        # 2) The card's summary line claims more files than we grabbed directly
        # (e.g. "PDF | 17 KB | PDF | 1 MB" but the actual links are hidden behind
        # numbered chip buttons like "PDF 1" / "PDF 2") -> click each missing one
        if ok_count < file_count:
            ok_count += download_numbered_chips(context, page, index, title, folder, used, ok_count, file_count)

        if ok_count >= file_count and ok_count > 0:
            return True, ""
        if ok_count > 0:
            print(f"  Note: only {ok_count}/{file_count} files downloaded for '{title}'")
            return True, ""
        reason = last_reason or "no downloadable link or chip found"

        # 3) Last resort -> click the whole card
        kind, result = click_card(context, page, index)
        if kind == "download":
            ext = Path(result.suggested_filename).suffix or ".pdf"
            path = unique_in_run(folder, title + ext, used)
            result.save_as(path)
            print(f"  Downloaded: {path.name}")
            return True, ""

        if kind == "url":
            response = context.request.get(result, timeout=TIMEOUT)
            content_type = response.headers.get("content-type", "")
            if not response.ok:
                reason = f"HTTP {response.status} for {result}"
            elif "text/html" in content_type:
                reason = f"opened a web page, not a file ({result})"
            else:
                path = unique_in_run(folder, title + extension_for(result, content_type), used)
                path.write_bytes(response.body())
                print(f"  Downloaded: {path.name}")
                return True, ""
        else:
            reason = "nothing happened after clicking"
    except Exception as e:
        reason = str(e).splitlines()[0] if str(e) else "error"

    print(f"  NOT downloaded: {title} ({reason})")
    return False, reason


def download_cards(context, page, url, folder, used):
    seen_before = set()  # cards from earlier pages (so "load more" lists are not repeated)
    done, failed = 0, []
    page_no = 1

    while page_no <= MAX_PAGES:
        load_all_cards(page)
        cards = page.evaluate(CARD_SCRIPT)

        if SAVE_DEBUG_HTML and page_no == 1:
            Path("debug_page.html").write_text(page.content(), encoding="utf-8")

        new_cards = [c for c in cards if (c["title"], tuple(c.get("hrefs", []))) not in seen_before]
        print(f"{url} -> page {page_no}: {len(cards)} cards found, {len(new_cards)} new")
        for n, c in enumerate(cards, 1):
            print(f"   {n}. {c['title'][:90]}")
        if not new_cards:
            break

        this_page = set()
        for i in range(len(cards)):
            current = page.evaluate(CARD_SCRIPT)  # re-tag in case the page redrew
            if i >= len(current):
                break
            card = current[i]
            key = (card["title"], tuple(card.get("hrefs", [])))
            if key in seen_before:
                continue
            this_page.add(key)

            ok, reason = process_card(context, page, i, card, folder, used)
            if ok:
                done += 1
            else:
                failed.append((card["title"], reason))

        seen_before |= this_page
        if not go_to_next_page(page):
            break
        page_no += 1

    print(f"{url} -> downloaded {done}, not downloaded {len(failed)}")
    for title, reason in failed:
        print(f"   - {title[:80]}: {reason}")


# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------


def goto_with_retry(page, url, attempts=4, wait_seconds=20):
    """Try loading the page a few times, waiting between tries. Sites sometimes
    briefly block or rate-limit automated requests after repeated runs."""
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            page.goto(url, wait_until="domcontentloaded")
            return
        except Exception as e:
            last_error = e
            print(f"  Attempt {attempt}/{attempts} failed to load {url}: {e}")
            if attempt < attempts:
                print(f"  Waiting {wait_seconds}s before retrying...")
                page.wait_for_timeout(wait_seconds * 1000)
    raise last_error


def download_from_site(context, url):
    parsed = urlparse(url)
    site_folder = re.sub(r"[^A-Za-z0-9._-]+", "_", parsed.netloc + parsed.path).strip("_")
    folder = DOWNLOAD_DIR / (site_folder[:100] or "site")
    folder.mkdir(parents=True, exist_ok=True)
    used = set()

    page = context.new_page()
    page.set_default_timeout(TIMEOUT)
    try:
        goto_with_retry(page, url)
        wait_for_content(page)

        if page.evaluate(CARD_SCRIPT):
            download_cards(context, page, url, folder, used)
        else:
            links = get_file_links(page)
            if links:
                print(f"{url} -> {len(links)} file links found")
                download_normal_links(context, links, folder, used)
            else:
                item_links = get_item_links(page)
                if item_links:
                    print(f"{url} -> no direct file links; found {len(item_links)} item pages, opening each")
                    download_item_search_results(context, page, url, folder, used)
                else:
                    print(f"{url} -> 0 file links found")
    except Exception as e:
        print(f"Could not process {url}: {e}")
    finally:
        page.close()


def main():
    DOWNLOAD_DIR.mkdir(exist_ok=True)
    with sync_playwright() as p:
        launch_args = {"headless": HEADLESS}
        context_args = {
            "accept_downloads": True,
            "user_agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
            ),
            "viewport": {"width": 1366, "height": 800},
        }
        if SYSTEM_PROXY:
            print(f"Using system proxy: {SYSTEM_PROXY}")
            launch_args["proxy"] = {"server": SYSTEM_PROXY}
            context_args["proxy"] = {"server": SYSTEM_PROXY}

        browser = p.chromium.launch(**launch_args)
        context = browser.new_context(**context_args)
        for url in URLS:
            download_from_site(context, url)
        context.close()
        browser.close()
    print("Done.")


if __name__ == "__main__":
    main()
