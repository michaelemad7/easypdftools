"""EasyPDFTools web layer (FastAPI).

All document processing lives in pdftools.py. This file handles HTTP concerns:
upload validation, rate limiting, response headers, SEO pages and the OCR client.

Configuration (all optional, set as Cloud Run environment variables):
  SITE_URL                 canonical site URL           (default https://ezpdftool.com)
  MAX_FILE_MB              per-request upload limit     (default 30; Cloud Run HTTP/1 caps at 32 MiB)
  MAX_FILES                files per request            (default 20)
  MAX_CONCURRENT_JOBS      parallel conversions/instance (default 3)
  RATE_LIMIT_PER_MIN       API calls per IP per minute  (default 20)
  OCR_LIMIT_PER_HOUR       OCR calls per IP per hour    (default 10)
  XFF_HOPS                 which X-Forwarded-For entry from the right is the client (default 1;
                           use 2 if a Google load balancer sits in front of Cloud Run)
  GCP_VISION_API_KEY       optional; if unset the Cloud Run service account is used instead
  ENABLE_DOCS=1            expose /docs for debugging
"""
import asyncio
import base64
import html
import json
import logging
import os
import re
import time
import unicodedata
from collections import defaultdict, deque
from functools import lru_cache
from pathlib import Path
from urllib.parse import quote

import httpx
from fastapi import (
    APIRouter,
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
)
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

import pdftools as pt

logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"))
log = logging.getLogger("easypdftools")

BASE_DIR = Path(__file__).resolve().parent
SITE_URL = os.getenv("SITE_URL", "https://ezpdftool.com").rstrip("/")
MAX_FILE_BYTES = int(float(os.getenv("MAX_FILE_MB", "30")) * 1024 * 1024)
MAX_FILE_MB = MAX_FILE_BYTES / (1024 * 1024)
MAX_FILES = int(os.getenv("MAX_FILES", "20"))
MAX_CONCURRENT_JOBS = int(os.getenv("MAX_CONCURRENT_JOBS", "3"))
RATE_LIMIT_PER_MIN = int(os.getenv("RATE_LIMIT_PER_MIN", "20"))
OCR_LIMIT_PER_HOUR = int(os.getenv("OCR_LIMIT_PER_HOUR", "10"))
XFF_HOPS = max(1, int(os.getenv("XFF_HOPS", "1")))
VISION_API_KEY = os.getenv("GCP_VISION_API_KEY", "")

VISION_URL = "https://vision.googleapis.com/v1/images:annotate"
METADATA_TOKEN_URL = (
    "http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token"
)

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF = "application/pdf"
ZIP = "application/zip"

# slug -> (page title, meta description). Each slug gets its own crawlable URL.
DEFAULT_PAGES = {
    "pdf-to-word": {"title": "PDF to Word Converter", "meta_description": "Convert PDF to an editable Word (.docx) document online. Free, fast, no signup.", "core_tool": "pdf-to-word", "faqs": []},
    "pdf-to-excel": {"title": "PDF to Excel Converter", "meta_description": "Extract tables and text from a PDF into an Excel (.xlsx) spreadsheet. Free, no signup.", "core_tool": "pdf-to-excel", "faqs": []},
    "pdf-to-jpg": {"title": "PDF to JPG Converter", "meta_description": "Turn every page of a PDF into a high-quality JPG image. Free and fast.", "core_tool": "pdf-to-jpg", "faqs": []},
    "pdf-to-text": {"title": "PDF to Text Converter", "meta_description": "Extract plain text from a PDF file online. Free, no signup.", "core_tool": "pdf-to-text", "faqs": []},
    "word-to-pdf": {"title": "Word to PDF Converter", "meta_description": "Convert Word (.docx) documents to PDF online. Free and fast.", "core_tool": "word-to-pdf", "faqs": []},
    "excel-to-pdf": {"title": "Excel to PDF Converter", "meta_description": "Convert Excel (.xlsx) spreadsheets to PDF online. Free and fast.", "core_tool": "excel-to-pdf", "faqs": []},
    "jpg-to-pdf": {"title": "JPG to PDF Converter", "meta_description": "Combine JPG and PNG images into a single PDF. Free, no signup.", "core_tool": "jpg-to-pdf", "faqs": []},
    "merge-pdf": {"title": "Merge PDF Files", "meta_description": "Combine multiple PDF files into one document online. Free and fast.", "core_tool": "merge-pdf", "faqs": []},
    "split-pdf": {"title": "Split PDF", "meta_description": "Split a PDF into single pages or custom page ranges. Free and fast.", "core_tool": "split-pdf", "faqs": []},
    "compress-pdf": {"title": "Compress PDF", "meta_description": "Reduce PDF file size by recompressing images and streams. Free, no signup.", "core_tool": "compress-pdf", "faqs": []},
    "sign-pdf": {"title": "Sign PDF", "meta_description": "Draw or upload your signature and place it on any page of a PDF, right in your browser.", "core_tool": "sign-pdf", "faqs": []},
    "ocr-pdf": {"title": "Scan PDF to Word (OCR)", "meta_description": "Turn scanned PDFs into editable Word documents with OCR text recognition.", "core_tool": "ocr-pdf", "faqs": []},
}

def load_seo_pages() -> dict:
    try:
        path = BASE_DIR / "seo_pages.json"
        if path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
            pages = {}
            for item in data:
                slug = item.get("slug", "").strip("/")
                if slug:
                    pages[slug] = {
                        "title": item.get("title", ""),
                        "meta_description": item.get("meta_description", ""),
                        "core_tool": item.get("core_tool", slug),
                        "faqs": item.get("faqs", [])
                    }
            if pages:
                merged = dict(DEFAULT_PAGES)
                merged.update(pages)
                return merged
    except Exception as e:
        log.error(f"Failed to load SEO pages: {e}")
    return DEFAULT_PAGES

SEO_PAGES = load_seo_pages()

# --------------------------------------------------------------------------------------
# App setup
# --------------------------------------------------------------------------------------

docs_on = os.getenv("ENABLE_DOCS") == "1"
app = FastAPI(
    title="EasyPDFTools",
    docs_url="/docs" if docs_on else None,
    redoc_url=None,
    openapi_url="/openapi.json" if docs_on else None,
)
app.mount("/static", StaticFiles(directory=BASE_DIR / "static", check_dir=False), name="static")

_jobs = asyncio.Semaphore(MAX_CONCURRENT_JOBS)


async def run(fn, *args):
    """Run blocking document work in a thread, a few at a time, without blocking the event loop."""
    async with _jobs:
        return await asyncio.to_thread(fn, *args)


def client_ip(request: Request) -> str:
    """Cloud Run appends the real client IP to X-Forwarded-For; earlier entries are spoofable."""
    forwarded = request.headers.get("x-forwarded-for", "")
    parts = [p.strip() for p in forwarded.split(",") if p.strip()]
    if parts:
        return parts[-min(XFF_HOPS, len(parts))]
    return request.client.host if request.client else "unknown"


class RateLimiter:
    """Sliding-window limiter, per instance. Good enough to stop casual abuse;
    use Cloud Armor or a shared store if you need a global limit."""

    def __init__(self, limit: int, window_seconds: int):
        self.limit = limit
        self.window = window_seconds
        self.hits: dict[str, deque] = defaultdict(deque)
        self._last_prune = time.monotonic()

    async def __call__(self, request: Request) -> None:
        now = time.monotonic()
        queue = self.hits[client_ip(request)]
        while queue and now - queue[0] > self.window:
            queue.popleft()
        if len(queue) >= self.limit:
            retry = int(self.window - (now - queue[0])) + 1
            raise HTTPException(
                status_code=429,
                detail="Too many requests. Please wait a moment and try again.",
                headers={"Retry-After": str(retry)},
            )
        queue.append(now)
        if now - self._last_prune > 60:
            self._last_prune = now
            for ip in [ip for ip, q in self.hits.items() if not q or now - q[-1] > self.window]:
                del self.hits[ip]


general_limiter = RateLimiter(RATE_LIMIT_PER_MIN, 60)
ocr_limiter = RateLimiter(OCR_LIMIT_PER_HOUR, 3600)
api = APIRouter(prefix="/api", dependencies=[Depends(general_limiter)])


# --------------------------------------------------------------------------------------
# Middleware and error handlers
# --------------------------------------------------------------------------------------


@app.middleware("http")
async def common_headers(request: Request, call_next):
    if request.method == "POST" and request.url.path.startswith("/api/"):
        length = request.headers.get("content-length", "")
        if length.isdigit() and int(length) > MAX_FILE_BYTES + 1_048_576:
            return JSONResponse(
                {"detail": f"The upload is larger than the {MAX_FILE_MB:.0f} MB limit."},
                status_code=413,
            )
    response = await call_next(request)
    headers = response.headers
    headers.setdefault("X-Content-Type-Options", "nosniff")
    headers.setdefault("X-Frame-Options", "DENY")
    headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    if request.headers.get("x-forwarded-proto") == "https":
        headers.setdefault("Strict-Transport-Security", "max-age=31536000")
    if request.url.path.startswith("/api/"):
        headers["Cache-Control"] = "no-store"
    # Content-Security-Policy is added together with the new front end (it must match
    # the scripts/styles the page loads).
    return response


@app.exception_handler(pt.ToolError)
async def tool_error_handler(request: Request, exc: pt.ToolError):
    return JSONResponse({"detail": exc.message}, status_code=exc.status)


@app.exception_handler(RequestValidationError)
async def validation_error_handler(request: Request, exc: RequestValidationError):
    return JSONResponse(
        {"detail": "The upload was missing or invalid. Please choose a file and try again."},
        status_code=422,
    )


@app.exception_handler(Exception)
async def unexpected_error_handler(request: Request, exc: Exception):
    return JSONResponse(
        {"detail": "Something went wrong while processing your file. Please try a different file."},
        status_code=500,
    )


# --------------------------------------------------------------------------------------
# Upload + response helpers
# --------------------------------------------------------------------------------------


def _display_name(file: UploadFile) -> str:
    name = re.sub(r"[\x00-\x1f\x7f]", "", file.filename or "file")
    return name if len(name) <= 60 else name[:57] + "..."


async def _read_limited(file: UploadFile, limit: int, label: str) -> bytes:
    chunks: list[bytes] = []
    size = 0
    while chunk := await file.read(1024 * 1024):
        size += len(chunk)
        if size > limit:
            raise pt.ToolError(label, 413)
        chunks.append(chunk)
    if not size:
        raise pt.ToolError(f'"{_display_name(file)}" is empty.', 400)
    return b"".join(chunks)


async def read_upload(file: UploadFile) -> bytes:
    return await _read_limited(
        file, MAX_FILE_BYTES, f'"{_display_name(file)}" is larger than the {MAX_FILE_MB:.0f} MB limit.'
    )


async def read_uploads(files: list[UploadFile]) -> list[bytes]:
    if len(files) > MAX_FILES:
        raise pt.ToolError(f"Please upload at most {MAX_FILES} files at a time.", 400)
    remaining = MAX_FILE_BYTES
    out = []
    for file in files:
        data = await _read_limited(
            file, remaining, f"The files together are larger than the {MAX_FILE_MB:.0f} MB limit."
        )
        remaining -= len(data)
        out.append(data)
    return out


def attachment(original: str | None, ext: str, suffix: str = "") -> dict[str, str]:
    """Content-Disposition that survives spaces, quotes and non-ASCII (e.g. Arabic) filenames."""
    base = re.split(r"[\\/]", original or "")[-1]
    stem = unicodedata.normalize("NFC", Path(base).stem)
    stem = re.sub(r'[\x00-\x1f\x7f/\\:*?"<>|]+', "_", stem).strip(" ._")[:80] or "document"
    ascii_stem = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode()
    ascii_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", ascii_stem).strip("_.") or "document"
    return {
        "Content-Disposition": (
            f'attachment; filename="{ascii_stem}{suffix}{ext}"; '
            f"filename*=UTF-8''{quote(stem + suffix + ext, safe='')}"
        )
    }


def send(data: bytes, media_type: str, original: str | None, ext: str, suffix: str = "",
         extra: dict[str, str] | None = None) -> Response:
    headers = attachment(original, ext, suffix)
    if extra:
        headers.update(extra)
    return Response(content=data, media_type=media_type, headers=headers)


# --------------------------------------------------------------------------------------
# Conversion endpoints
# --------------------------------------------------------------------------------------


@api.post("/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    data = await read_upload(file)
    return send(await run(pt.pdf_to_docx, data), DOCX, file.filename, ".docx")


@api.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    data = await read_upload(file)
    return send(await run(pt.pdf_to_xlsx, data), XLSX, file.filename, ".xlsx")


@api.post("/pdf-to-jpg")
async def pdf_to_jpg(file: UploadFile = File(...)):
    data = await read_upload(file)
    return send(await run(pt.pdf_to_jpg_zip, data), ZIP, file.filename, ".zip", suffix="_images")


@api.post("/pdf-to-text")
async def pdf_to_text(file: UploadFile = File(...)):
    data = await read_upload(file)
    return send(await run(pt.pdf_to_txt, data), "text/plain; charset=utf-8", file.filename, ".txt")


@api.post("/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    data = await read_upload(file)
    return send(await run(pt.docx_to_pdf, data), PDF, file.filename, ".pdf")


@api.post("/excel-to-pdf")
async def excel_to_pdf(file: UploadFile = File(...)):
    data = await read_upload(file)
    return send(await run(pt.xlsx_to_pdf, data), PDF, file.filename, ".pdf")


@api.post("/jpg-to-pdf")
async def jpg_to_pdf(files: list[UploadFile] = File(...)):
    datas = await read_uploads(files)
    return send(await run(pt.images_to_pdf, datas), PDF, "images", ".pdf")


@api.post("/merge")
async def merge(files: list[UploadFile] = File(...)):
    if len(files) < 2:
        raise pt.ToolError("Upload at least 2 PDFs to merge.", 400)
    datas = await read_uploads(files)
    return send(await run(pt.merge_pdfs, datas), PDF, "merged", ".pdf")


@api.post("/split")
async def split(file: UploadFile = File(...), ranges: str = Form("")):
    data = await read_upload(file)
    if len(ranges) > 200:
        raise pt.ToolError("The page range is too long.", 400)
    return send(await run(pt.split_pdf, data, ranges), ZIP, file.filename, ".zip", suffix="_split")


@api.post("/reduce")
async def reduce(file: UploadFile = File(...), level: str = Form("medium")):
    data = await read_upload(file)
    result, before, after = await run(pt.reduce_pdf, data, level)
    return send(
        result, PDF, file.filename, ".pdf", suffix="_compressed",
        extra={"X-Original-Size": str(before), "X-Result-Size": str(after)},
    )


# --------------------------------------------------------------------------------------
# OCR (Google Cloud Vision)
# --------------------------------------------------------------------------------------

_token_cache = {"value": "", "expires": 0.0}


async def vision_headers() -> dict[str, str]:
    """API key if configured, otherwise the Cloud Run service account's access token."""
    if VISION_API_KEY:
        return {"X-Goog-Api-Key": VISION_API_KEY}
    now = time.time()
    if _token_cache["value"] and now < _token_cache["expires"] - 60:
        return {"Authorization": f"Bearer {_token_cache['value']}"}
    try:
        async with httpx.AsyncClient(timeout=5.0) as client:
            res = await client.get(METADATA_TOKEN_URL, headers={"Metadata-Flavor": "Google"})
            res.raise_for_status()
            payload = res.json()
    except Exception:
        log.exception("Could not get a service-account token for the Vision API")
        raise pt.ToolError("OCR isn't available on this server right now.", 503)
    _token_cache["value"] = payload["access_token"]
    _token_cache["expires"] = now + float(payload.get("expires_in", 300))
    return {"Authorization": f"Bearer {_token_cache['value']}"}


async def vision_page(client: httpx.AsyncClient, headers: dict, jpeg: bytes, hints: list[str]) -> list[str]:
    request = {
        "image": {"content": base64.b64encode(jpeg).decode("ascii")},
        "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
    }
    if hints:
        request["imageContext"] = {"languageHints": hints}
    try:
        res = await client.post(VISION_URL, json={"requests": [request]}, headers=headers)
    except httpx.HTTPError:
        log.exception("Vision API request failed")
        raise pt.ToolError("The OCR service couldn't be reached. Please try again.", 503)
    if res.status_code != 200:
        log.error("Vision API returned %s: %s", res.status_code, res.text[:500])
        raise pt.ToolError("The OCR service returned an error. Please try again later.", 502)
    first = (res.json().get("responses") or [{}])[0]
    if "error" in first:
        log.error("Vision API error: %s", first["error"])
        raise pt.ToolError("The OCR service couldn't process this page.", 502)
    return pt.vision_paragraphs(first.get("fullTextAnnotation") or {})


@api.post("/edit", dependencies=[Depends(ocr_limiter)])
async def edit_ocr(file: UploadFile = File(...), languages: str = Form("")):
    """Scan a PDF to an editable Word file (first OCR_MAX_PAGES pages)."""
    data = await read_upload(file)
    hints = [h for h in re.split(r"[,\s]+", languages) if re.fullmatch(r"[A-Za-z]{2,3}(-[A-Za-z]{2,4})?", h)][:5]
    images, total = await run(pt.pdf_to_jpeg_pages, data, pt.OCR_MAX_PAGES)
    headers = await vision_headers()
    async with httpx.AsyncClient(timeout=httpx.Timeout(60.0, connect=10.0)) as client:
        pages = await asyncio.gather(*(vision_page(client, headers, img, hints) for img in images))
    docx_bytes = await run(pt.build_ocr_docx, list(pages))
    return send(
        docx_bytes, DOCX, file.filename, ".docx", suffix="_ocr",
        extra={"X-Pages-Processed": str(len(images)), "X-Pages-Total": str(total)},
    )


app.include_router(api)


# --------------------------------------------------------------------------------------
# Pages, health check and SEO files
# --------------------------------------------------------------------------------------

def _build_faq_ld_and_html(faqs: list[dict]) -> tuple[str, str]:
    if not faqs:
        return "", ""
    main_entities = []
    html_items = ['<div class="faq-container"><h2>Frequently Asked Questions</h2>']
    for faq in faqs:
        q = (faq.get("question") or faq.get("q") or "").strip()
        a = (faq.get("answer") or faq.get("a") or "").strip()
        if not q or not a:
            continue
        main_entities.append({
            "@type": "Question",
            "name": q,
            "acceptedAnswer": {"@type": "Answer", "text": a}
        })
        html_items.append(f'<div class="faq-item"><h3>{html.escape(q)}</h3><p>{html.escape(a)}</p></div>')
    
    html_items.append("</div>")
    if not main_entities:
        return "", ""

    ld_struct = {
        "@context": "https://schema.org",
        "@type": "FAQPage",
        "mainEntity": main_entities
    }
    ld_script = f'<script type="application/ld+json">{json.dumps(ld_struct)}</script>'
    return ld_script, "\n".join(html_items)

@lru_cache(maxsize=None)
def _template() -> str:
    return (BASE_DIR / "templates" / "index.html").read_text(encoding="utf-8")

@lru_cache(maxsize=None)
def render_page(slug: str = "") -> str:
    """Serve index.html with a per-tool <title>, description and canonical URL."""
    page = _template()
    slug_clean = slug.strip("/")
    page_info = SEO_PAGES.get(slug_clean)
    
    if page_info:
        title = page_info.get("title", "")
        description = page_info.get("meta_description", "")
        core_tool = page_info.get("core_tool", slug_clean)
        faqs = page_info.get("faqs", [])
        
        if title:
            page = re.sub(r"<title>.*?</title>", lambda _: f"<title>{html.escape(title)} | EasyPDFTools</title>", page, count=1, flags=re.S)
        if description:
            page = re.sub(r'(<meta\s+name="description"\s+content=")[^"]*(")', lambda m: m.group(1) + html.escape(description, quote=True) + m.group(2), page, count=1)
            
        faq_ld, faq_html = _build_faq_ld_and_html(faqs)
        
        head_injections = []
        if core_tool:
            head_injections.append(f'<meta name="core-tool" content="{html.escape(core_tool, quote=True)}">')
        if faq_ld:
            head_injections.append(faq_ld)
        if 'rel="canonical"' not in page:
            canonical_url = f"{SITE_URL}/{slug_clean}" if slug_clean else SITE_URL
            head_injections.append(f'<link rel="canonical" href="{canonical_url}">')
            
        if head_injections:
            page = page.replace("</head>", "\n".join(head_injections) + "\n</head>", 1)
            
        if faq_html:
            page = re.sub(r'(<section\s+id="faq-section"[^>]*>)(.*?)(</section>)', rf"\1\n{faq_html}\n\3", page, count=1, flags=re.S)
            
    else:
        if 'rel="canonical"' not in page:
            canonical = f'<link rel="canonical" href="{SITE_URL}/{slug_clean}">'
            page = page.replace("</head>", f"{canonical}\n</head>", 1)
            
    return page

def _page_response(slug: str = "") -> HTMLResponse:
    return HTMLResponse(render_page(slug), headers={"Cache-Control": "public, max-age=300"})

@app.get("/", include_in_schema=False)
async def index():
    return _page_response()

def _register_tool_page(slug: str) -> None:
    async def tool_page():
        return _page_response(slug)

    app.add_api_route(f"/{slug}", tool_page, methods=["GET"], include_in_schema=False)

for _slug in SEO_PAGES:
    _register_tool_page(_slug)

@app.get("/healthz", include_in_schema=False)
async def healthz():
    return {"status": "ok"}

FAVICON_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
    '<rect width="100" height="100" rx="25" fill="#e11d48"/>'
    '<text x="50%" y="55%" dominant-baseline="central" text-anchor="middle" fill="white" '
    'font-family="sans-serif" font-weight="900" font-size="50">EP</text></svg>'
)

@app.get("/favicon.ico", include_in_schema=False)
@app.get("/favicon.svg", include_in_schema=False)
async def favicon():
    return Response(FAVICON_SVG, media_type="image/svg+xml", headers={"Cache-Control": "public, max-age=86400"})

@app.get("/robots.txt", include_in_schema=False)
async def robots_txt():
    body = f"User-agent: *\nAllow: /\nDisallow: /api/\nSitemap: {SITE_URL}/sitemap.xml\n"
    return Response(body, media_type="text/plain")

@app.get("/ads.txt", include_in_schema=False)
async def ads_txt():
    return Response("google.com, pub-2658744954536274, DIRECT, f08c47fec0942fa0\n", media_type="text/plain")

@app.get("/sitemap.xml", include_in_schema=False)
async def sitemap_xml():
    urls = [f"  <url><loc>{SITE_URL}/</loc><priority>1.0</priority></url>"]
    urls += [f"  <url><loc>{SITE_URL}/{slug}</loc><priority>0.8</priority></url>" for slug in SEO_PAGES]
    body = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + "\n".join(urls) + "\n</urlset>"
    )
    return Response(body, media_type="application/xml")