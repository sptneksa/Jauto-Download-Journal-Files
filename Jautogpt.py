from __future__ import annotations

import hashlib
import logging
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Mapping
from urllib.parse import quote, unquote, urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ============================================================
# KONFIGURASI UTAMA
# ============================================================

INPUT_EXCEL = Path("Jauto.xlsx")
OUTPUT_EXCEL = Path("Jauto_Status_OpenAccess.xlsx")
OUTPUT_FOLDER = Path("Jurnal_Unduhan")
LOG_FILE = Path("download_open_access.log")

# Disarankan menggunakan environment variable.
USER_EMAIL = os.getenv(
    "SCHOLAR_EMAIL",
    "ekaseptian354@gmail.com"
).strip()

OPENALEX_API_KEY = os.getenv(
    "OPENALEX_API_KEY",
    ""
).strip()

# API Key Elsevier dari user. Tetap bisa dioverride lewat environment variable
# agar tidak perlu menyimpan API key langsung di kode ketika dipakai di mesin lain.
ELSEVIER_API_KEY = os.getenv(
    "ELSEVIER_API_KEY",
    "115ed441cd8c6f76d430d30b337d9761"
).strip()

DOAJ_API_BASE = "https://doaj.org/api/search/articles"

# Timeout terdiri dari:
# (waktu koneksi, waktu membaca respons)
REQUEST_TIMEOUT = (15, 60)

# Batas maksimal ukuran PDF.
MAX_PDF_SIZE_MB = 150

# Jeda antarbaris untuk menghindari terlalu banyak permintaan.
PAUSE_BETWEEN_ROWS = 1.0

# Kolom yang wajib tersedia dalam Excel.
REQUIRED_COLUMNS = {
    "Sub",
    "Title",
    "DOI",
}

# Pola standar DOI.
DOI_PATTERN = re.compile(
    r"10\.\d{4,9}/[-._;()/:A-Z0-9]+",
    re.IGNORECASE,
)


# ============================================================
# MODEL DATA KANDIDAT URL
# ============================================================

@dataclass(frozen=True)
class Candidate:
    """
    Menyimpan kandidat URL artikel.

    kind:
        pdf     = URL diduga langsung menuju berkas PDF.
        landing = URL menuju halaman artikel/repository.

    extra_headers:
        Header tambahan untuk provider tertentu, misalnya Elsevier API
        yang membutuhkan X-ELS-APIKey. Disimpan sebagai tuple agar aman
        untuk dataclass frozen.
    """

    url: str
    source: str
    kind: str = "pdf"
    extra_headers: tuple[tuple[str, str], ...] = ()


# ============================================================
# LOGGING
# ============================================================

def setup_logging() -> None:
    """
    Mengaktifkan log ke terminal dan file.
    """

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(
                LOG_FILE,
                encoding="utf-8",
            ),
        ],
    )


# ============================================================
# SESSION HTTP DENGAN RETRY
# ============================================================

def build_session() -> requests.Session:
    """
    Membuat requests.Session dengan mekanisme retry otomatis.
    """

    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        backoff_factor=1.2,
        status_forcelist=(
            429,
            500,
            502,
            503,
            504,
        ),
        allowed_methods=frozenset({
            "GET",
            "HEAD",
        }),
        respect_retry_after_header=True,
        raise_on_status=False,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=10,
        pool_maxsize=10,
    )

    session = requests.Session()

    session.mount(
        "https://",
        adapter,
    )

    session.mount(
        "http://",
        adapter,
    )

    session.headers.update({
        "User-Agent": (
            f"OpenAccessArticleDownloader/3.0 "
            f"(mailto:{USER_EMAIL})"
        ),
        "Accept-Language": "en-US,en;q=0.9,id;q=0.8",
    })

    return session


# ============================================================
# NORMALISASI DOI
# ============================================================

def normalize_doi(value: object) -> str | None:
    """
    Membersihkan DOI dari berbagai format.

    Contoh input:
        doi:10.1234/abcd
        https://doi.org/10.1234/abcd
        http://dx.doi.org/10.1234/abcd

    Output:
        10.1234/abcd
    """

    if pd.isna(value):
        return None

    text = unquote(str(value)).strip()

    text = re.sub(
        r"^(?:https?://(?:dx\.)?doi\.org/|doi\s*:\s*)",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = text.strip().strip(
        "\"'<>[]{}"
    )

    match = DOI_PATTERN.search(text)

    if not match:
        return None

    return match.group(0).rstrip(
        ".,;"
    ).lower()


# ============================================================
# PEMBERSIHAN NAMA FILE
# ============================================================

def safe_filename_component(
    value: object,
    max_length: int = 110,
) -> str:
    """
    Membersihkan karakter yang tidak diperbolehkan
    dalam nama file Windows.
    """

    text = str(value).strip()

    text = re.sub(
        r'[\\/*?:"<>|]',
        "",
        text,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip(" .")

    if not text:
        text = "tanpa-judul"

    return text[:max_length].rstrip(" .")


# ============================================================
# MENGHAPUS URL DUPLIKAT
# ============================================================

def unique_candidates(
    candidates: Iterable[Candidate],
) -> list[Candidate]:
    """
    Menghapus kandidat URL yang sama.
    """

    result: list[Candidate] = []
    seen: set[str] = set()

    for candidate in candidates:
        url = str(
            candidate.url or ""
        ).strip()

        if not url.startswith((
            "http://",
            "https://",
        )):
            continue

        normalized_url = url.split(
            "#",
            1,
        )[0]

        if normalized_url in seen:
            continue

        seen.add(normalized_url)

        result.append(
            Candidate(
                url=normalized_url,
                source=candidate.source,
                kind=candidate.kind,
                extra_headers=candidate.extra_headers,
            )
        )

    return result


# ============================================================
# MENYIMPAN PROGRES EXCEL
# ============================================================

def save_progress(df: pd.DataFrame) -> None:
    """
    Menyimpan progres secara atomik agar file Excel
    tidak mudah rusak saat program berhenti.
    """

    temp_path = OUTPUT_EXCEL.with_name(
        f"{OUTPUT_EXCEL.stem}.tmp.xlsx"
    )

    df.to_excel(
        temp_path,
        index=False,
    )

    os.replace(
        temp_path,
        OUTPUT_EXCEL,
    )


# ============================================================
# VALIDASI FORMAT PDF
# ============================================================

def looks_like_pdf(data: bytes) -> bool:
    """
    Memeriksa signature PDF.

    PDF normal memiliki teks %PDF- pada bagian awal file.
    """

    return b"%PDF-" in data[:1024]


# ============================================================
# MENGUNDUH PDF
# ============================================================

def download_pdf(
    session: requests.Session,
    pdf_url: str,
    save_path: Path,
    extra_headers: Mapping[str, str] | None = None,
) -> tuple[bool, str, str]:
    """
    Mengunduh dan memvalidasi PDF.

    Return:
        success
        pesan
        final_url
    """

    temp_path = save_path.with_suffix(
        save_path.suffix + ".part"
    )

    max_bytes = (
        MAX_PDF_SIZE_MB
        * 1024
        * 1024
    )

    request_headers = {
        "Accept": "application/pdf,*/*;q=0.8",
    }

    if extra_headers:
        request_headers.update(dict(extra_headers))

    try:
        with session.get(
            pdf_url,
            stream=True,
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
            headers=request_headers,
        ) as response:

            final_url = response.url

            if response.status_code != 200:
                return (
                    False,
                    f"HTTP {response.status_code}",
                    final_url,
                )

            content_length = response.headers.get(
                "Content-Length"
            )

            if (
                content_length
                and content_length.isdigit()
                and int(content_length) > max_bytes
            ):
                return (
                    False,
                    (
                        f"Ukuran file melebihi "
                        f"{MAX_PDF_SIZE_MB} MB"
                    ),
                    final_url,
                )

            chunks = response.iter_content(
                chunk_size=64 * 1024
            )

            first_chunk = next(
                chunks,
                b"",
            )

            if not first_chunk:
                return (
                    False,
                    "Respons kosong",
                    final_url,
                )

            if not looks_like_pdf(first_chunk):
                content_type = response.headers.get(
                    "Content-Type",
                    "",
                ).lower()

                return (
                    False,
                    (
                        "Bukan PDF "
                        f"({content_type or 'content-type kosong'})"
                    ),
                    final_url,
                )

            total_size = len(first_chunk)

            save_path.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            with open(
                temp_path,
                "wb",
            ) as file:
                file.write(first_chunk)

                for chunk in chunks:
                    if not chunk:
                        continue

                    total_size += len(chunk)

                    if total_size > max_bytes:
                        raise ValueError(
                            (
                                "Ukuran file melebihi "
                                f"{MAX_PDF_SIZE_MB} MB"
                            )
                        )

                    file.write(chunk)

        if total_size < 1024:
            temp_path.unlink(
                missing_ok=True
            )

            return (
                False,
                "PDF terlalu kecil atau rusak",
                final_url,
            )

        os.replace(
            temp_path,
            save_path,
        )

        return (
            True,
            "OK",
            final_url,
        )

    except (
        requests.RequestException,
        OSError,
        ValueError,
    ) as exc:

        temp_path.unlink(
            missing_ok=True
        )

        return (
            False,
            f"{type(exc).__name__}: {exc}",
            pdf_url,
        )


# ============================================================
# MENCARI PDF DARI LANDING PAGE
# ============================================================

def extract_pdf_links_from_landing(
    session: requests.Session,
    landing_url: str,
    source: str,
) -> list[Candidate]:
    """
    Membuka landing page artikel dan mencari URL PDF publik
    melalui metadata HTML.
    """

    try:
        response = session.get(
            landing_url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            headers={
                "Accept": (
                    "text/html,"
                    "application/xhtml+xml,"
                    "application/pdf;q=0.9,"
                    "*/*;q=0.8"
                ),
            },
        )

        if response.status_code != 200:
            return []

        raw_content = response.content

        # Apabila respons langsung berupa PDF.
        if looks_like_pdf(raw_content):
            return [
                Candidate(
                    url=response.url,
                    source=source,
                    kind="pdf",
                )
            ]

        content_type = response.headers.get(
            "Content-Type",
            "",
        ).lower()

        if (
            "html" not in content_type
            and b"<html" not in raw_content[:2000].lower()
        ):
            return []

        soup = BeautifulSoup(
            raw_content,
            "html.parser",
        )

        found: list[Candidate] = []

        # Metadata yang umum digunakan website jurnal.
        meta_names = {
            "citation_pdf_url",
            "eprints.document_url",
            "wkhealth_pdf_url",
            "bepress_citation_pdf_url",
        }

        for meta in soup.find_all("meta"):
            name = str(
                meta.get("name")
                or meta.get("property")
                or ""
            ).lower()

            content = str(
                meta.get("content")
                or ""
            ).strip()

            if (
                name in meta_names
                and content
            ):
                found.append(
                    Candidate(
                        url=urljoin(
                            response.url,
                            content,
                        ),
                        source=source,
                        kind="pdf",
                    )
                )

        # Tag link dengan content-type PDF.
        for tag in soup.select(
            'link[type="application/pdf"][href]'
        ):
            found.append(
                Candidate(
                    url=urljoin(
                        response.url,
                        tag["href"],
                    ),
                    source=source,
                    kind="pdf",
                )
            )

        # Embed, iframe, dan object.
        for tag in soup.select(
            "embed[src], iframe[src], object[data]"
        ):
            value = (
                tag.get("src")
                or tag.get("data")
            )

            if not value:
                continue

            value = str(value)

            if (
                "pdf" in value.lower()
                or value.lower().endswith(".pdf")
            ):
                found.append(
                    Candidate(
                        url=urljoin(
                            response.url,
                            value,
                        ),
                        source=source,
                        kind="pdf",
                    )
                )

        # Fallback: tautan yang berakhiran .pdf.
        for tag in soup.select("a[href]"):
            href = str(
                tag.get("href")
                or ""
            ).strip()

            clean_href = href.split(
                "?",
                1,
            )[0].lower()

            if clean_href.endswith(".pdf"):
                found.append(
                    Candidate(
                        url=urljoin(
                            response.url,
                            href,
                        ),
                        source=source,
                        kind="pdf",
                    )
                )

        return unique_candidates(
            found
        )[:10]

    except requests.RequestException as exc:
        logging.debug(
            "Gagal membuka landing page %s: %s",
            landing_url,
            exc,
        )

        return []


# ============================================================
# SUMBER 1: OPENALEX
# ============================================================

def candidates_from_openalex(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Mengambil kandidat PDF dari OpenAlex.
    """

    encoded_doi = quote(
        doi,
        safe="",
    )

    url = (
        "https://api.openalex.org/works/"
        f"https://doi.org/{encoded_doi}"
    )

    params: dict[str, str] = {}

    if USER_EMAIL:
        params["mailto"] = USER_EMAIL

    if OPENALEX_API_KEY:
        params["api_key"] = OPENALEX_API_KEY

    try:
        response = session.get(
            url,
            params=params,
            timeout=REQUEST_TIMEOUT,
        )

        if response.status_code != 200:
            return (
                [],
                f"OpenAlex HTTP {response.status_code}",
            )

        data = response.json()
        locations: list[dict] = []

        # Prioritaskan lokasi OA terbaik dan lokasi utama.
        for key in (
            "best_oa_location",
            "primary_location",
        ):
            location = data.get(key)

            if isinstance(
                location,
                dict,
            ):
                locations.append(location)

        # Tambahkan seluruh lokasi lain.
        for location in data.get(
            "locations"
        ) or []:
            if isinstance(
                location,
                dict,
            ):
                locations.append(location)

        candidates: list[Candidate] = []

        for location in locations:
            # Abaikan lokasi yang bukan open access.
            if location.get("is_oa") is not True:
                continue

            pdf_url = location.get(
                "pdf_url"
            )

            landing_page_url = location.get(
                "landing_page_url"
            )

            if pdf_url:
                candidates.append(
                    Candidate(
                        url=pdf_url,
                        source="OpenAlex",
                        kind="pdf",
                    )
                )

            if landing_page_url:
                candidates.append(
                    Candidate(
                        url=landing_page_url,
                        source="OpenAlex",
                        kind="landing",
                    )
                )

        return (
            unique_candidates(candidates),
            "OK",
        )

    except (
        requests.RequestException,
        ValueError,
    ) as exc:
        return (
            [],
            (
                f"OpenAlex "
                f"{type(exc).__name__}: {exc}"
            ),
        )


# ============================================================
# SUMBER 2: UNPAYWALL
# ============================================================

def candidates_from_unpaywall(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Mengambil kandidat PDF dari Unpaywall.
    """

    if (
        not USER_EMAIL
        or USER_EMAIL == "emailanda@gmail.com"
    ):
        return (
            [],
            (
                "Unpaywall dilewati: "
                "SCHOLAR_EMAIL belum diatur"
            ),
        )

    encoded_doi = quote(
        doi,
        safe="/",
    )

    url = (
        "https://api.unpaywall.org/v2/"
        f"{encoded_doi}"
    )

    try:
        response = session.get(
            url,
            params={
                "email": USER_EMAIL,
            },
            timeout=REQUEST_TIMEOUT,
        )

        if response.status_code != 200:
            return (
                [],
                f"Unpaywall HTTP {response.status_code}",
            )

        data = response.json()
        locations: list[dict] = []

        best_location = data.get(
            "best_oa_location"
        )

        if isinstance(
            best_location,
            dict,
        ):
            locations.append(best_location)

        for location in data.get(
            "oa_locations"
        ) or []:
            if isinstance(
                location,
                dict,
            ):
                locations.append(location)

        candidates: list[Candidate] = []

        for location in locations:
            pdf_url = location.get(
                "url_for_pdf"
            )

            landing_url = location.get(
                "url_for_landing_page"
            )

            if pdf_url:
                candidates.append(
                    Candidate(
                        url=pdf_url,
                        source="Unpaywall",
                        kind="pdf",
                    )
                )

            if landing_url:
                candidates.append(
                    Candidate(
                        url=landing_url,
                        source="Unpaywall",
                        kind="landing",
                    )
                )

        return (
            unique_candidates(candidates),
            "OK",
        )

    except (
        requests.RequestException,
        ValueError,
    ) as exc:
        return (
            [],
            (
                f"Unpaywall "
                f"{type(exc).__name__}: {exc}"
            ),
        )


# ============================================================
# SUMBER 3: DOAJ API
# ============================================================

def _bibjson_matches_doi(
    bibjson: dict,
    doi: str,
) -> bool:
    """
    Memastikan record DOAJ benar-benar sesuai dengan DOI target.
    """

    for identifier in bibjson.get("identifier") or []:
        if not isinstance(identifier, dict):
            continue

        identifier_type = str(
            identifier.get("type") or ""
        ).lower()

        identifier_value = normalize_doi(
            identifier.get("id")
        )

        if (
            identifier_type == "doi"
            and identifier_value == doi
        ):
            return True

    return False


def _doaj_query_candidates(
    session: requests.Session,
    query: str,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Menjalankan satu query DOAJ dan mengekstrak link full-text/PDF.
    """

    url = f"{DOAJ_API_BASE}/{quote(query, safe='')}"

    response = session.get(
        url,
        params={
            "page": "1",
            "pageSize": "5",
        },
        timeout=REQUEST_TIMEOUT,
        headers={
            "Accept": "application/json",
        },
    )

    if response.status_code != 200:
        return (
            [],
            f"DOAJ HTTP {response.status_code}",
        )

    data = response.json()
    results = data.get("results") or []

    candidates: list[Candidate] = []

    for result in results:
        if not isinstance(result, dict):
            continue

        bibjson = result.get("bibjson") or {}

        if not isinstance(bibjson, dict):
            continue

        if not _bibjson_matches_doi(bibjson, doi):
            # Query DOAJ kadang longgar; hindari salah unduh.
            continue

        for link in bibjson.get("link") or []:
            if not isinstance(link, dict):
                continue

            link_url = str(
                link.get("url") or ""
            ).strip()

            if not link_url.startswith((
                "http://",
                "https://",
            )):
                continue

            link_type = str(
                link.get("type") or ""
            ).lower()

            content_type = str(
                link.get("content_type")
                or link.get("content-type")
                or ""
            ).lower()

            clean_url = link_url.split(
                "?",
                1,
            )[0].lower()

            if (
                clean_url.endswith(".pdf")
                or "pdf" in content_type
            ):
                kind = "pdf"
            else:
                kind = "landing"

            # DOAJ biasanya menyimpan fulltext URL, bisa landing atau PDF.
            if (
                "fulltext" in link_type
                or kind == "pdf"
            ):
                candidates.append(
                    Candidate(
                        url=link_url,
                        source="DOAJ",
                        kind=kind,
                    )
                )

    return (
        unique_candidates(candidates),
        "OK" if candidates else "DOAJ: tidak ada full-text URL untuk DOI",
    )


def candidates_from_doaj(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Mengambil kandidat artikel/PDF dari DOAJ Public Search API.
    
    DOAJ tidak selalu menyediakan URL PDF langsung; sering kali yang tersedia
    adalah full-text landing page. Karena itu, landing page tetap diperiksa
    lagi oleh extract_pdf_links_from_landing().
    """

    queries = [
        f'bibjson.identifier.id:"{doi}"',
        f'doi:"{doi}"',
        f'"{doi}"',
    ]

    notes: list[str] = []

    for query in queries:
        try:
            candidates, note = _doaj_query_candidates(
                session=session,
                query=query,
                doi=doi,
            )

            if candidates:
                return (
                    candidates,
                    "OK",
                )

            if note != "OK":
                notes.append(note)

        except (
            requests.RequestException,
            ValueError,
        ) as exc:
            notes.append(
                (
                    "DOAJ "
                    f"{type(exc).__name__}: {exc}"
                )
            )

    return (
        [],
        " | ".join(notes[-3:]) or "DOAJ: tidak ditemukan",
    )


# ============================================================
# SUMBER 4: ELSEVIER ARTICLE RETRIEVAL API
# ============================================================

def _truthy_elsevier_flag(value: object) -> bool:
    """
    Membaca variasi flag open-access dari metadata Elsevier.
    """

    return str(value or "").strip().lower() in {
        "1",
        "true",
        "full",
        "open",
        "yes",
    }


def _extract_elsevier_coredata(data: dict) -> dict:
    """
    Mengambil coredata dari respons JSON Elsevier yang formatnya nested.
    """

    response = data.get("full-text-retrieval-response") or {}

    if isinstance(response, dict):
        coredata = response.get("coredata") or {}
        if isinstance(coredata, dict):
            return coredata

    coredata = data.get("coredata") or {}
    return coredata if isinstance(coredata, dict) else {}


def candidates_from_elsevier(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Mengambil PDF melalui Elsevier Article Retrieval API.

    Provider ini hanya ditambahkan sebagai kandidat bila metadata Elsevier
    menunjukkan artikel open-access. Jika API key tidak memiliki entitlement
    atau artikel bukan OA, provider akan dilewati secara aman.
    """

    if not ELSEVIER_API_KEY:
        return (
            [],
            "Elsevier dilewati: ELSEVIER_API_KEY belum diatur",
        )

    encoded_doi = quote(
        doi,
        safe="",
    )

    base_url = (
        "https://api.elsevier.com/content/article/doi/"
        f"{encoded_doi}"
    )

    headers = {
        "X-ELS-APIKey": ELSEVIER_API_KEY,
        "Accept": "application/json",
    }

    try:
        response = session.get(
            base_url,
            params={
                "httpAccept": "application/json",
            },
            headers=headers,
            timeout=REQUEST_TIMEOUT,
        )

        if response.status_code != 200:
            return (
                [],
                f"Elsevier HTTP {response.status_code}",
            )

        data = response.json()
        coredata = _extract_elsevier_coredata(data)

        is_open_access = any(
            _truthy_elsevier_flag(coredata.get(key))
            for key in (
                "openaccess",
                "openaccessArticle",
                "openaccessType",
            )
        )

        if not is_open_access:
            return (
                [],
                "Elsevier: metadata tidak menunjukkan open-access",
            )

        pdf_url = (
            f"{base_url}?httpAccept=application/pdf"
        )

        return (
            [
                Candidate(
                    url=pdf_url,
                    source="Elsevier API",
                    kind="pdf",
                    extra_headers=(
                        ("X-ELS-APIKey", ELSEVIER_API_KEY),
                        ("Accept", "application/pdf"),
                    ),
                )
            ],
            "OK",
        )

    except (
        requests.RequestException,
        ValueError,
    ) as exc:
        return (
            [],
            (
                "Elsevier "
                f"{type(exc).__name__}: {exc}"
            ),
        )


# ============================================================
# SUMBER 5: SCI-HUB / BLACK OPEN ACCESS
# ============================================================

def candidates_from_scihub_disabled(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Sci-Hub tidak diimplementasikan.

    Program ini hanya mengunduh PDF dari sumber open-access/legal
    atau API resmi. Otomasi Sci-Hub berisiko melanggar hak cipta,
    sehingga provider ini sengaja dibuat nonaktif.
    """

    return (
        [],
        "Sci-Hub dilewati: gunakan sumber open-access/legal atau API resmi",
    )


# ============================================================
# SUMBER 3: SEMANTIC SCHOLAR
# ============================================================

def candidates_from_semantic_scholar(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Mengambil openAccessPdf dari Semantic Scholar.
    """

    paper_id = quote(
        f"DOI:{doi}",
        safe="",
    )

    url = (
        "https://api.semanticscholar.org/"
        f"graph/v1/paper/{paper_id}"
    )

    headers: dict[str, str] = {}

    if SEMANTIC_SCHOLAR_API_KEY:
        headers["x-api-key"] = (
            SEMANTIC_SCHOLAR_API_KEY
        )

    try:
        response = session.get(
            url,
            params={
                "fields": "title,openAccessPdf",
            },
            headers=headers,
            timeout=REQUEST_TIMEOUT,
        )

        if response.status_code != 200:
            return (
                [],
                (
                    "Semantic Scholar HTTP "
                    f"{response.status_code}"
                ),
            )

        data = response.json()

        open_access_pdf = (
            data.get("openAccessPdf")
            or {}
        )

        if isinstance(
            open_access_pdf,
            dict,
        ):
            pdf_url = open_access_pdf.get(
                "url"
            )
        else:
            pdf_url = None

        if pdf_url:
            return (
                [
                    Candidate(
                        url=pdf_url,
                        source="Semantic Scholar",
                        kind="pdf",
                    )
                ],
                "OK",
            )

        return (
            [],
            (
                "Semantic Scholar: "
                "openAccessPdf kosong"
            ),
        )

    except (
        requests.RequestException,
        ValueError,
    ) as exc:
        return (
            [],
            (
                "Semantic Scholar "
                f"{type(exc).__name__}: {exc}"
            ),
        )


# ============================================================
# SUMBER 4: CROSSREF
# ============================================================

def candidates_from_crossref(
    session: requests.Session,
    doi: str,
) -> tuple[list[Candidate], str]:
    """
    Mengambil tautan PDF yang tercantum dalam metadata Crossref.
    """

    encoded_doi = quote(
        doi,
        safe="",
    )

    url = (
        "https://api.crossref.org/works/"
        f"{encoded_doi}"
    )

    try:
        response = session.get(
            url,
            params={
                "mailto": USER_EMAIL,
            },
            timeout=REQUEST_TIMEOUT,
        )

        if response.status_code != 200:
            return (
                [],
                f"Crossref HTTP {response.status_code}",
            )

        message = (
            response.json().get("message")
            or {}
        )

        candidates: list[Candidate] = []

        for link in message.get(
            "link"
        ) or []:
            if (
                not isinstance(link, dict)
                or not link.get("URL")
            ):
                continue

            content_type = str(
                link.get("content-type")
                or ""
            ).lower()

            link_url = str(
                link["URL"]
            )

            clean_url = link_url.split(
                "?",
                1,
            )[0].lower()

            if (
                "pdf" in content_type
                or clean_url.endswith(".pdf")
            ):
                candidates.append(
                    Candidate(
                        url=link_url,
                        source="Crossref",
                        kind="pdf",
                    )
                )

        return (
            unique_candidates(candidates),
            "OK",
        )

    except (
        requests.RequestException,
        ValueError,
    ) as exc:
        return (
            [],
            (
                f"Crossref "
                f"{type(exc).__name__}: {exc}"
            ),
        )


# ============================================================
# FALLBACK LANDING PAGE DOI
# ============================================================

def doi_landing_candidate(
    doi: str,
) -> list[Candidate]:
    """
    Membuat kandidat landing page melalui doi.org.
    """

    encoded_doi = quote(
        doi,
        safe="/",
    )

    return [
        Candidate(
            url=f"https://doi.org/{encoded_doi}",
            source="DOI Landing Page",
            kind="landing",
        )
    ]


# ============================================================
# DAFTAR PROVIDER
# ============================================================

PROVIDERS: list[
    Callable[
        [requests.Session, str],
        tuple[list[Candidate], str],
    ]
] = [
    # Prioritas utama sesuai kebutuhan revisi: sumber legal/open-access.
    candidates_from_unpaywall,
    candidates_from_openalex,
    candidates_from_doaj,
    candidates_from_elsevier,

    # Sengaja nonaktif; hanya memberi catatan pada output/log.
    candidates_from_scihub_disabled,

    # Fallback tambahan legal. Boleh dihapus bila hanya ingin 4 provider utama.
    candidates_from_crossref,
]


# ============================================================
# PROSES UTAMA
# ============================================================

def download_open_access_articles(
    excel_file: Path = INPUT_EXCEL,
) -> None:
    """
    Membaca daftar DOI dari Excel dan mencari PDF open-access.
    """

    setup_logging()

    OUTPUT_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Resume dari file hasil jika sudah tersedia.
    if OUTPUT_EXCEL.exists():
        source_excel = OUTPUT_EXCEL
        logging.info(
            "Melanjutkan proses dari file status: %s",
            OUTPUT_EXCEL,
        )
    else:
        source_excel = excel_file

    if not source_excel.exists():
        logging.error(
            "File Excel tidak ditemukan: %s",
            source_excel,
        )
        return

    try:
        df = pd.read_excel(
            source_excel
        )

    except Exception as exc:
        logging.error(
            "Gagal membaca file Excel: %s",
            exc,
        )
        return

    # Memeriksa kolom wajib.
    missing_columns = REQUIRED_COLUMNS.difference(
        df.columns
    )

    if missing_columns:
        logging.error(
            "Kolom wajib tidak ditemukan: %s",
            ", ".join(
                sorted(missing_columns)
            ),
        )
        return

    # Menambahkan kolom laporan jika belum ada.
    default_columns = {
        "Status Download": "Belum diproses",
        "Sumber Download": "",
        "URL PDF": "",
        "Catatan": "",
        "Waktu Proses": "",
    }

    for column, default_value in default_columns.items():
        if column not in df.columns:
            df[column] = default_value

    # Peringatan konfigurasi.
    if USER_EMAIL == "emailanda@gmail.com":
        logging.warning(
            "Email masih menggunakan placeholder. "
            "Ganti USER_EMAIL atau atur environment variable "
            "SCHOLAR_EMAIL. Unpaywall akan dilewati."
        )

    if not OPENALEX_API_KEY:
        logging.warning(
            "OPENALEX_API_KEY belum diatur. "
            "Permintaan OpenAlex tetap dapat berjalan, tetapi "
            "dapat terkena pembatasan jika skala besar."
        )

    if not ELSEVIER_API_KEY:
        logging.warning(
            "ELSEVIER_API_KEY belum diatur. "
            "Provider Elsevier API akan dilewati."
        )

    session = build_session()
    total_rows = len(df)

    for index, row in df.iterrows():
        sub = safe_filename_component(
            row.get("Sub", ""),
            max_length=40,
        )

        title = safe_filename_component(
            row.get("Title", ""),
            max_length=120,
        )

        doi = normalize_doi(
            row.get("DOI")
        )

        # Hash membuat nama file tetap unik.
        doi_or_index = doi or str(index)

        doi_hash = hashlib.sha1(
            doi_or_index.encode("utf-8")
        ).hexdigest()[:8]

        file_name = (
            f"{sub}-{title}-{doi_hash}.pdf"
        )

        save_path = (
            OUTPUT_FOLDER
            / file_name
        )

        logging.info(
            "[%s/%s] Memproses: %s",
            index + 1,
            total_rows,
            title,
        )

        # Jika file sudah ada dan ukurannya masuk akal.
        if (
            save_path.exists()
            and save_path.stat().st_size >= 1024
        ):
            logging.info(
                "File sudah tersedia: %s",
                save_path,
            )

            df.at[
                index,
                "Status Download"
            ] = "Berhasil - file sudah ada"

            df.at[
                index,
                "Catatan"
            ] = (
                "Dilewati karena PDF valid "
                "sudah tersedia"
            )

            df.at[
                index,
                "Waktu Proses"
            ] = datetime.now().isoformat(
                timespec="seconds"
            )

            save_progress(df)
            continue

        # DOI kosong atau tidak valid.
        if not doi:
            logging.warning(
                "DOI kosong atau tidak valid."
            )

            df.at[
                index,
                "Status Download"
            ] = "Gagal - DOI tidak valid/kosong"

            df.at[
                index,
                "Catatan"
            ] = (
                "DOI tidak cocok dengan "
                "pola DOI standar"
            )

            df.at[
                index,
                "Waktu Proses"
            ] = datetime.now().isoformat(
                timespec="seconds"
            )

            save_progress(df)
            continue

        logging.info(
            "DOI: %s",
            doi,
        )

        all_candidates: list[Candidate] = []
        provider_notes: list[str] = []

        # Memanggil seluruh provider.
        for provider in PROVIDERS:
            try:
                candidates, note = provider(
                    session,
                    doi,
                )

                all_candidates.extend(
                    candidates
                )

                if note != "OK":
                    provider_notes.append(
                        note
                    )

            except Exception as exc:
                provider_notes.append(
                    (
                        f"{provider.__name__}: "
                        f"{type(exc).__name__}: {exc}"
                    )
                )

        # Menambahkan fallback melalui DOI landing page.
        all_candidates.extend(
            doi_landing_candidate(doi)
        )

        all_candidates = unique_candidates(
            all_candidates
        )

        downloaded = False
        attempts: list[str] = []

        logging.info(
            "Ditemukan %s kandidat URL.",
            len(all_candidates),
        )

        for candidate in all_candidates:
            candidates_to_download: list[Candidate]

            if candidate.kind == "landing":
                logging.info(
                    "Memeriksa landing page dari %s",
                    candidate.source,
                )

                candidates_to_download = (
                    extract_pdf_links_from_landing(
                        session=session,
                        landing_url=candidate.url,
                        source=candidate.source,
                    )
                )

                if not candidates_to_download:
                    attempts.append(
                        (
                            f"{candidate.source}: "
                            "landing page tanpa PDF publik"
                        )
                    )
                    continue

            else:
                candidates_to_download = [
                    candidate
                ]

            for pdf_candidate in candidates_to_download:
                logging.info(
                    "Mencoba PDF dari %s",
                    pdf_candidate.source,
                )

                success, message, final_url = download_pdf(
                    session=session,
                    pdf_url=pdf_candidate.url,
                    save_path=save_path,
                    extra_headers=dict(pdf_candidate.extra_headers),
                )

                if success:
                    downloaded = True

                    df.at[
                        index,
                        "Status Download"
                    ] = "Berhasil"

                    df.at[
                        index,
                        "Sumber Download"
                    ] = pdf_candidate.source

                    df.at[
                        index,
                        "URL PDF"
                    ] = final_url

                    df.at[
                        index,
                        "Catatan"
                    ] = (
                        "PDF open-access berhasil "
                        "diunduh dan divalidasi"
                    )

                    logging.info(
                        "Berhasil melalui %s",
                        pdf_candidate.source,
                    )

                    break

                attempts.append(
                    (
                        f"{pdf_candidate.source}: "
                        f"{message}"
                    )
                )

                logging.debug(
                    "Gagal dari %s: %s",
                    pdf_candidate.source,
                    message,
                )

            if downloaded:
                break

        # Apabila seluruh sumber gagal.
        if not downloaded:
            df.at[
                index,
                "Status Download"
            ] = "Tidak ditemukan PDF open-access"

            df.at[
                index,
                "Sumber Download"
            ] = ""

            df.at[
                index,
                "URL PDF"
            ] = ""

            combined_notes = (
                provider_notes
                + attempts
            )

            # Membatasi panjang catatan Excel.
            df.at[
                index,
                "Catatan"
            ] = " | ".join(
                combined_notes[-12:]
            )[:3000]

            logging.warning(
                (
                    "PDF open-access tidak ditemukan "
                    "untuk DOI %s"
                ),
                doi,
            )

        df.at[
            index,
            "Waktu Proses"
        ] = datetime.now().isoformat(
            timespec="seconds"
        )

        # Simpan setelah setiap artikel agar progres tidak hilang.
        save_progress(df)

        time.sleep(
            PAUSE_BETWEEN_ROWS
        )

    logging.info(
        "Proses selesai."
    )

    logging.info(
        "Status tersimpan di: %s",
        OUTPUT_EXCEL,
    )

    logging.info(
        "Folder PDF: %s",
        OUTPUT_FOLDER,
    )


# ============================================================
# MENJALANKAN PROGRAM
# ============================================================

if __name__ == "__main__":
    download_open_access_articles(
        INPUT_EXCEL
    )