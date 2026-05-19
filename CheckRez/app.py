from __future__ import annotations

import io
import json
import os
import re
import sqlite3
import time
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import datetime
from typing import List
from urllib.parse import parse_qs, unquote, urlparse

import requests
from bs4 import BeautifulSoup
from docx import Document
from flask import Flask, Response, redirect, render_template, request, session, url_for
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
app.secret_key = os.getenv("FLASK_SECRET_KEY", "checkrez-dev-secret-change-me")
DB_PATH = "checkrez_history.db"
RATE_BUCKETS: dict[str, deque[float]] = defaultdict(deque)
RATE_LIMIT = 30
RATE_WINDOW_SEC = 60


def load_local_env(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except Exception:
        pass


load_local_env()


@dataclass
class SearchResult:
    title: str
    snippet: str
    link: str
    source: str = ""


@dataclass
class VacancyHints:
    title: str = ""
    company: str = ""
    city: str = ""


@dataclass
class CompanyProfile:
    name: str = ""
    inn: str = ""
    ogrn: str = ""
    kpp: str = ""
    status: str = ""
    address: str = ""
    ceo: str = ""
    source: str = ""


def init_db() -> None:
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS checks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mode TEXT NOT NULL,
            source_url TEXT,
            score INTEGER NOT NULL,
            entities TEXT,
            categories TEXT,
            facts_true TEXT,
            facts_doubt TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    try:
        cur.execute("ALTER TABLE checks ADD COLUMN user_id INTEGER")
    except sqlite3.OperationalError:
        pass
    con.commit()
    con.close()


def save_check(mode: str, source_url: str, score: int, entities: list[str], categories: dict, truths: list[str], doubts: list[str]) -> int:
    uid = session.get("user_id")
    if not uid:
        return 0
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(
        """
        INSERT INTO checks (mode, source_url, score, entities, categories, facts_true, facts_doubt, created_at, user_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            mode,
            source_url,
            score,
            json.dumps(entities, ensure_ascii=False),
            json.dumps(categories, ensure_ascii=False),
            json.dumps(truths, ensure_ascii=False),
            json.dumps(doubts, ensure_ascii=False),
            datetime.utcnow().isoformat(timespec="seconds"),
            uid,
        ),
    )
    check_id = cur.lastrowid
    con.commit()
    con.close()
    return int(check_id)


def get_check(check_id: int):
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    row = con.execute("SELECT * FROM checks WHERE id = ?", (check_id,)).fetchone()
    con.close()
    return row


def get_recent_checks(limit: int = 25):
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    uid = session.get("user_id")
    if uid:
        rows = con.execute("SELECT * FROM checks WHERE user_id = ? ORDER BY id DESC LIMIT ?", (uid, limit)).fetchall()
    else:
        rows = []
    con.close()
    return rows


def clear_user_history(user_id: int) -> None:
    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM checks WHERE user_id = ?", (user_id,))
    con.commit()
    con.close()


def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    row = con.execute("SELECT id, username, email FROM users WHERE id = ?", (uid,)).fetchone()
    con.close()
    return row


def enforce_rate_limit() -> str | None:
    ip = (request.headers.get("X-Forwarded-For") or request.remote_addr or "anon").split(",")[0].strip()
    now = time.time()
    q = RATE_BUCKETS[ip]
    while q and now - q[0] > RATE_WINDOW_SEC:
        q.popleft()
    if len(q) >= RATE_LIMIT:
        return "Слишком много запросов. Подожди минуту и повтори проверку."
    q.append(now)
    return None


def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    return "\n".join((page.extract_text() or "") for page in reader.pages).strip()


def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in doc.paragraphs if p.text).strip()


def fetch_url_text(url: str) -> tuple[str, BeautifulSoup | None, List[str]]:
    warnings: List[str] = []
    if not url:
        return "", None, warnings
    try:
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"}:
            raise ValueError
        r = requests.get(url, timeout=12, headers={"User-Agent": "Mozilla/5.0"})
        if not r.ok:
            warnings.append("Не удалось загрузить текст по ссылке.")
            return "", None, warnings
        soup = BeautifulSoup(r.text, "html.parser")
        text = " ".join(soup.stripped_strings)[:15000]
        return text, soup, warnings
    except Exception:
        warnings.append("Ссылка недоступна или некорректна.")
        return "", None, warnings


def extract_input_text(raw_text: str, source_url: str, uploaded_file) -> tuple[str, List[str], BeautifulSoup | None]:
    warnings = []
    text = (raw_text or "").strip()
    soup = None

    if uploaded_file and uploaded_file.filename:
        file_bytes = uploaded_file.read()
        filename = uploaded_file.filename.lower()
        try:
            if filename.endswith(".pdf"):
                text = extract_text_from_pdf(file_bytes)
            elif filename.endswith(".docx"):
                text = extract_text_from_docx(file_bytes)
            elif filename.endswith(".doc"):
                warnings.append("Формат .doc поддерживается ограниченно. Лучше .docx или PDF.")
            else:
                warnings.append("Неподдерживаемый формат файла. Используй PDF или DOCX.")
        except Exception:
            warnings.append("Не удалось прочитать файл. Попробуй другой документ.")

    if not text and source_url:
        text, soup, url_warnings = fetch_url_text(source_url)
        warnings.extend(url_warnings)

    return text, warnings, soup


def extract_hh_vacancy_hints(url: str, soup: BeautifulSoup | None) -> VacancyHints:
    if "hh.ru/vacancy/" not in url and "headhunter.ru/vacancy/" not in url:
        return VacancyHints()
    if not soup:
        return VacancyHints()

    title = (soup.select_one("h1[data-qa='vacancy-title']") or soup.select_one("h1"))
    company = soup.select_one("[data-qa='vacancy-company-name']")
    city = soup.select_one("[data-qa='vacancy-view-raw-address']")
    return VacancyHints(
        title=title.get_text(" ", strip=True) if title else "",
        company=company.get_text(" ", strip=True) if company else "",
        city=city.get_text(" ", strip=True) if city else "",
    )


def normalize_result_link(link: str) -> str:
    if not link:
        return ""
    parsed = urlparse(link)
    if "duckduckgo.com" in parsed.netloc and parsed.path.startswith("/l/"):
        target = parse_qs(parsed.query).get("uddg", [""])[0]
        return unquote(target) if target else link
    return link


def ddg_search(query: str, limit: int = 8) -> List[SearchResult]:
    if not query:
        return []
    try:
        r = requests.post(
            "https://html.duckduckgo.com/html/",
            data={"q": query},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=12,
        )
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        out = []
        for node in soup.select("div.result")[:limit]:
            a = node.select_one("a.result__a")
            snippet = node.select_one("a.result__snippet") or node.select_one("div.result__snippet")
            if not a:
                continue
            link = normalize_result_link(a.get("href", ""))
            host = (urlparse(link).netloc or "").lower().replace("www.", "")
            out.append(SearchResult(a.get_text(" ", strip=True), snippet.get_text(" ", strip=True) if snippet else "", link, host))
        return out
    except Exception:
        return []


def bing_search(query: str, limit: int = 8) -> List[SearchResult]:
    if not query:
        return []
    try:
        r = requests.get(
            "https://www.bing.com/search",
            params={"q": query, "count": str(limit)},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=12,
        )
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        out: List[SearchResult] = []
        for li in soup.select("li.b_algo")[:limit]:
            a = li.select_one("h2 a")
            p = li.select_one("p")
            if not a:
                continue
            link = a.get("href", "")
            host = (urlparse(link).netloc or "").lower().replace("www.", "")
            out.append(SearchResult(a.get_text(" ", strip=True), p.get_text(" ", strip=True) if p else "", link, host))
        return out
    except Exception:
        return []


def brave_search(query: str, limit: int = 8) -> List[SearchResult]:
    if not query:
        return []
    try:
        r = requests.get(
            "https://search.brave.com/search",
            params={"q": query, "source": "web"},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=12,
        )
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        out: List[SearchResult] = []
        for node in soup.select("div.snippet, div.fdb")[: limit * 2]:
            a = node.select_one("a[href]")
            if not a:
                continue
            title = a.get_text(" ", strip=True)
            link = a.get("href", "")
            sn = node.get_text(" ", strip=True)
            host = (urlparse(link).netloc or "").lower().replace("www.", "")
            if title and link.startswith("http"):
                out.append(SearchResult(title, sn[:220], link, host))
            if len(out) >= limit:
                break
        return out
    except Exception:
        return []


def unique_results(results: List[SearchResult], limit: int = 20) -> List[SearchResult]:
    seen = set()
    cleaned = []
    for r in results:
        key = r.link.lower().strip()
        if not key or key in seen:
            continue
        seen.add(key)
        cleaned.append(r)
        if len(cleaned) >= limit:
            break
    return cleaned


def normalize_company_token(company: str) -> str:
    token = re.sub(r"[^A-Za-zА-Яа-я0-9 ]+", " ", (company or "").lower())
    token = re.sub(r"\b(ооо|ао|зао|ип|llc|ltd|inc|company|corp|пао)\b", " ", token)
    token = re.sub(r"\s+", " ", token).strip()
    return token


def filter_employee_review_results(results: List[SearchResult], company: str) -> List[SearchResult]:
    review_markers = [
        "отзывы сотрудников",
        "отзывы работников",
        "работодатель",
        "conditions",
        "employer review",
        "glassdoor",
        "indeed",
        "dreamjob",
        "правда сотрудников",
    ]
    bad_markers = ["значение слова", "википедия", "dictionary", "перевод", "what is", "словарь"]
    company_token = normalize_company_token(company)
    keep: List[SearchResult] = []
    for r in results:
        text = f"{r.title} {r.snippet} {r.source}".lower()
        if any(x in text for x in bad_markers):
            continue
        has_review = any(x in text for x in review_markers)
        has_company = not company_token or any(tok in text for tok in company_token.split() if len(tok) > 2)
        if has_review and has_company:
            keep.append(r)
    return keep


def score_result_quality(r: SearchResult, context: str) -> int:
    s = 0
    host = r.source.lower()
    tt = f"{r.title} {r.snippet}".lower()
    trusted = ["linkedin.com", "hh.ru", "superjob.ru", "zarplata.ru", "glassdoor", "indeed", "spark-interfax", "rusprofile", "egrul"]
    review = ["otzovik", "dreamjob", "pravda-sotrudnikov", "glassdoor", "indeed"]
    noisy = ["pinterest", "tiktok", "youtube.com/shorts"]

    if any(x in host for x in trusted):
        s += 4
    if any(x in host for x in review):
        s += 3
    if any(x in host for x in noisy):
        s -= 4
    for token in context.lower().split()[:8]:
        if token and token in tt:
            s += 1
    return s


def search_web_multi(queries: List[str], deep: bool = False) -> tuple[List[SearchResult], dict[str, int]]:
    per_query = 10 if deep else 6
    bag = []
    for q in [x.strip() for x in queries if x and x.strip()]:
        res = ddg_search(q, per_query)
        if len(res) < 2:
            res.extend(bing_search(q, per_query))
        if len(res) < 2:
            res.extend(brave_search(q, per_query))
        bag.extend(res)
    bag = unique_results(bag, 100)
    ranked = sorted(bag, key=lambda x: score_result_quality(x, " ".join(queries)), reverse=True)
    top = ranked[:20 if deep else 12]

    weights = {
        "Профили/карьера": sum(1 for x in top if any(d in x.source for d in ["linkedin", "hh.ru", "superjob", "zarplata"])),
        "Отзывы": sum(1 for x in top if any(d in x.source for d in ["otzovik", "dreamjob", "glassdoor", "indeed", "pravda-sotrudnikov"])),
        "Реестры/юридическое": sum(1 for x in top if any(d in x.source for d in ["egrul", "rusprofile", "spark-interfax", "nalog"])),
        "Прочие": 0,
    }
    weights["Прочие"] = max(0, len(top) - sum(weights.values()))
    return top, weights


def detect_name(text: str) -> str:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines[:8]:
        if 2 <= len(ln.split()) <= 4 and len(ln) < 60:
            return ln
    m = re.search(r"([A-ZА-Я][a-zа-я]+\s+[A-ZА-Я][a-zа-я]+(?:\s+[A-ZА-Я][a-zа-я]+)?)", text)
    return m.group(1) if m else ""


def detect_companies(text: str) -> List[str]:
    pattern = r"(?:ООО|АО|ЗАО|ИП|LLC|Ltd|Inc|Company|Corp)\s+[\w\-\"'«»]{2,}"
    return list(dict.fromkeys(re.findall(pattern, text, flags=re.IGNORECASE)))[:5]


def detect_inn(text: str) -> str:
    m = re.search(r"\b\d{10}\b|\b\d{12}\b", text)
    return m.group(0) if m else ""


def infer_inn_from_web(company_name: str) -> str:
    if not company_name:
        return ""
    try:
        results = ddg_search(f"{company_name} ИНН ОГРН", limit=6)
        bag = " ".join([f"{r.title} {r.snippet}" for r in results])
        m = re.search(r"\b\d{10}\b|\b\d{12}\b", bag)
        return m.group(0) if m else ""
    except Exception:
        return ""


def fetch_company_profile(company_name: str = "", inn: str = "") -> CompanyProfile | None:
    token = os.getenv("DADATA_TOKEN", "").strip()
    if not token:
        return None
    query = (inn or company_name).strip()
    if not query:
        return None
    try:
        by_id = bool(re.fullmatch(r"\d{10}|\d{12}|\d{13}|\d{15}", query))
        endpoint = (
            "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
            if by_id
            else "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/party"
        )
        resp = requests.post(
            endpoint,
            headers={
                "Authorization": f"Token {token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            json={"query": query, "count": 1},
            timeout=12,
        )
        resp.raise_for_status()
        items = (resp.json() or {}).get("suggestions", [])
        if not items:
            return None
        d = items[0].get("data", {})
        return CompanyProfile(
            name=items[0].get("value", "") or company_name,
            inn=d.get("inn", ""),
            ogrn=d.get("ogrn", ""),
            kpp=d.get("kpp", ""),
            status=(d.get("state") or {}).get("status", ""),
            address=(d.get("address") or {}).get("value", ""),
            ceo=((d.get("management") or {}).get("name", "") + " " + (d.get("management") or {}).get("post", "")).strip(),
            source="DaData API",
        )
    except Exception:
        return None


def detect_salary_flags(text: str) -> tuple[list[str], list[str], int, list[str]]:
    truths, doubts, explain = [], [], []
    penalty = 0
    t = text.lower()
    if re.search(r"\d[\d\s]{3,}\s*(₽|руб|rub)", t):
        truths.append("В тексте есть внятная зарплатная информация.")
        explain.append("-2: указана оплата")
        penalty -= 2
    if "без опыта" in t and any(x in t for x in ["300 000", "500 000", "миллион"]):
        penalty += 12
        doubts.append("Завышенный доход для позиции без опыта.")
        explain.append("+12: завышенные обещания")
    risky = ["ежедневные выплаты", "легкие деньги", "перевод на карту", "оплата обучения", "предоплата"]
    hits = [x for x in risky if x in t]
    if hits:
        penalty += 5 * len(hits)
        doubts.append("Есть красные флаги по условиям оплаты/найма.")
        explain.append(f"+{5 * len(hits)}: red flags ({', '.join(hits[:3])})")
    return truths, doubts, penalty, explain


def detect_contact_flags(text: str) -> tuple[list[str], list[str], int, list[str]]:
    truths, doubts, explain = [], [], []
    penalty = 0
    t = text.lower()
    has_email = bool(re.search(r"[\w\.-]+@[\w\.-]+\.[a-z]{2,}", t))
    has_phone = bool(re.search(r"(\+7|8)\s*\(?\d{3}\)?[\s-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}", text))
    if has_email or has_phone:
        truths.append("Найдены контактные данные.")
        explain.append("-2: контакты присутствуют")
        penalty -= 2
    else:
        doubts.append("Нет явных контактных данных.")
        explain.append("+6: отсутствуют контакты")
        penalty += 6
    if "telegram" in t or "whatsapp" in t:
        doubts.append("Контакт через мессенджер требует доп. проверки.")
        explain.append("+3: контакт через мессенджер")
        penalty += 3
    return truths, doubts, penalty, explain


def detect_doc_consistency(text: str) -> tuple[list[str], list[str], int, list[str]]:
    truths, doubts, explain = [], [], []
    penalty = 0
    years = re.findall(r"(19\d{2}|20\d{2})", text)
    if len(years) >= 2:
        truths.append("Обнаружены даты, можно проверить хронологию.")
        explain.append("-1: есть хронология")
        penalty -= 1
    else:
        doubts.append("Мало дат и фактов для проверки хронологии.")
        explain.append("+4: слабая хронология")
        penalty += 4
    return truths, doubts, penalty, explain


def build_category_scores(score: int, text: str, results_count: int, flags_count: int) -> dict[str, int]:
    return {
        "Детализация": max(1, min(100, 100 - len(text) // 85)),
        "Веб-подтверждение": max(1, min(100, 100 - results_count * 8)),
        "Риторический риск": max(1, min(100, score + (10 if "срочно" in text.lower() else 0))),
        "Красные флаги": max(1, min(100, 20 + flags_count * 12)),
    }


def risk_to_auth(value: int) -> int:
    return max(1, min(100, 101 - value))


def convert_categories_to_auth(risk_categories: dict[str, int]) -> dict[str, int]:
    return {
        "Детализация профиля": risk_to_auth(risk_categories.get("Детализация", 50)),
        "Веб-подтвержденность": risk_to_auth(risk_categories.get("Веб-подтверждение", 50)),
        "Надежность формулировок": risk_to_auth(risk_categories.get("Риторический риск", 50)),
        "Чистота по red flags": risk_to_auth(risk_categories.get("Красные флаги", 50)),
    }


def convert_explain_to_auth(explain_risk: list[str]) -> list[str]:
    out: list[str] = []
    for item in explain_risk:
        m = re.match(r"^([+-])(\d+):\s*(.+)$", item.strip())
        if not m:
            out.append(item)
            continue
        sign, num, text = m.groups()
        new_sign = "+" if sign == "-" else "-"
        out.append(f"{new_sign}{num}: {text}")
    return out


def build_core_score(text: str, results: List[SearchResult], entities: list[str]) -> tuple[int, list[str], list[str], list[str]]:
    score = 55
    truths: list[str] = []
    doubts: list[str] = []
    explain: list[str] = []

    if len(text) > 600:
        score -= 10
        truths.append("Документ достаточно подробный.")
        explain.append("-10: высокая детализация")
    else:
        score += 15
        doubts.append("Короткий текст, сложнее проверить.")
        explain.append("+15: низкая детализация")

    if results:
        delta = min(25, len(results) * 2)
        score -= delta
        truths.append(f"Найдены внешние источники: {len(results)}")
        explain.append(f"-{delta}: есть внешние подтверждения")
    else:
        score += 20
        doubts.append("Внешние упоминания почти не найдены.")
        explain.append("+20: нет веб-подтверждений")

    if entities:
        score -= 8
        truths.append("Выделены ключевые сущности для проверки.")
        explain.append("-8: выделены сущности")
    else:
        score += 8
        doubts.append("Сущности выделить не удалось.")
        explain.append("+8: нет сущностей")

    suspicious = ["гарантированно", "без опыта", "легкие деньги", "срочно", "100%"]
    hits = sum(1 for w in suspicious if w in text.lower())
    if hits:
        score += hits * 4
        doubts.append("Есть сомнительные маркетинговые формулировки.")
        explain.append(f"+{hits * 4}: подозрительная риторика")

    return max(1, min(100, score)), truths, doubts, explain


def analyze_resume(text: str, deep: bool = False):
    name = detect_name(text)
    companies = detect_companies(text)
    base = " ".join([x for x in [name] + companies[:2] if x]).strip()
    chunk = " ".join(text.split())[:140]
    queries = [chunk]
    if base:
        queries = [f"{base} опыт работы", f"{base} linkedin", f"{base} резюме"] + queries

    results, source_weights = search_web_multi(queries, deep=deep)
    entities = [x for x in [name] + companies if x]

    score, truths, doubts, explain = build_core_score(text, results, entities)
    mods = [detect_contact_flags(text), detect_doc_consistency(text)]
    flags_count = 0
    for t, d, p, e in mods:
        truths.extend(t)
        doubts.extend(d)
        explain.extend(e)
        score = max(1, min(100, score + p))
        flags_count += len(d)

    categories = convert_categories_to_auth(build_category_scores(score, text, len(results), flags_count))
    score = risk_to_auth(score)
    checks = [
        "Сверить ФИО и опыт с LinkedIn/портфолио.",
        "Попросить 1-2 подтверждающих кейса по ключевым проектам.",
        "Проверить последовательность дат в опыте и обучении.",
    ]
    if deep:
        checks.append("Deep-check: сделать ручную верификацию 3 самых весомых источников.")

    explain_auth = convert_explain_to_auth(explain)
    return score, truths, doubts, results, entities, categories, checks, explain_auth, source_weights


def analyze_vacancy(text: str, source_url: str = "", hh_hints: VacancyHints | None = None, deep: bool = False):
    companies = detect_companies(text)
    company = companies[0] if companies else ""
    title = hh_hints.title if hh_hints else ""
    city = hh_hints.city if hh_hints else ""
    if hh_hints and hh_hints.company:
        company = hh_hints.company

    inn = detect_inn(text)
    if company and title:
        queries = [
            f"{company} {title} отзывы сотрудников",
            f"site:glassdoor.com {company} employee reviews",
            f"site:indeed.com {company} reviews",
            f"site:dreamjob.ru {company} отзывы сотрудников",
            f"{company} работодатель отзывы сотрудников",
            f"{company} ИНН ОГРН",
        ]
    elif company:
        queries = [
            f"{company} отзывы сотрудников",
            f"site:glassdoor.com {company} employee reviews",
            f"site:dreamjob.ru {company} отзывы сотрудников",
            f"{company} работодатель отзывы сотрудников",
            f"{company} ИНН ОГРН",
        ]
    else:
        frag = " ".join(text.split())[:140]
        queries = [f"{frag} отзывы о работодателе", f"{frag} отзывы сотрудников компания"]

    results, source_weights = search_web_multi(queries, deep=deep)
    filtered = []
    for r in results:
        link_low = r.link.lower()
        if "duckduckgo.com" in link_low:
            continue
        if "hh.ru" in link_low and "vacancy/" not in link_low:
            continue
        filtered.append(r)
    review_quality_note = ""
    if company:
        strict_reviews = filter_employee_review_results(filtered, company)
        if strict_reviews:
            filtered = strict_reviews
            review_quality_note = "good"
        else:
            review_quality_note = "weak"

    entities = [x for x in [company, title, city] if x]
    score, truths, doubts, explain = build_core_score(text, filtered, entities)

    mods = [detect_salary_flags(text), detect_contact_flags(text), detect_doc_consistency(text)]
    flags_count = 0
    for t, d, p, e in mods:
        truths.extend(t)
        doubts.extend(d)
        explain.extend(e)
        score = max(1, min(100, score + p))
        flags_count += len(d)

    if company:
        truths.append(f"Компания: {company}")
    else:
        doubts.append("Не удалось надежно выделить компанию.")
    if review_quality_note == "good":
        truths.append("Найдены релевантные отзывы сотрудников именно по компании.")
    elif review_quality_note == "weak":
        doubts.append("Релевантные отзывы сотрудников по компании найдены слабо.")
    if title:
        truths.append(f"Позиция: {title}")
    if source_url and "hh.ru/vacancy/" in source_url:
        truths.append("Ссылка распознана как карточка вакансии HH с приоритетным парсингом.")

    categories = convert_categories_to_auth(build_category_scores(score, text, len(filtered), flags_count))
    checks = [
        "Сверить ИНН/ОГРН и статус компании в открытых реестрах.",
        "Проверить отзывы на независимых площадках и форумах.",
        "Уточнить юрлицо в договоре и схему выплат.",
    ]
    if deep:
        checks.append("Deep-check: ручной аудит 5 источников с максимальным весом доверия.")

    profile = fetch_company_profile(company, inn)
    if not profile and company and not inn:
        inferred_inn = infer_inn_from_web(company)
        if inferred_inn:
            profile = fetch_company_profile(company, inferred_inn)
            if profile:
                truths.append(f"ИНН найден через веб-источники: {inferred_inn}")
                explain.append("-2: найден ИНН по внешним источникам")
                score = max(1, score - 2)
    if profile:
        truths.append("Получена структурированная карточка компании из API.")
        explain.append("-4: подтверждение компании через API")
        score = max(1, score - 4)
    else:
        doubts.append("API-карточка компании недоступна (нет токена или не найдено совпадение).")
        explain.append("+2: нет структурированного API-подтверждения")
        score = min(100, score + 2)

    score = risk_to_auth(score)
    explain_auth = convert_explain_to_auth(explain)
    return score, truths, doubts, filtered, entities, categories, checks, explain_auth, source_weights, profile


def render_pdf_report(check_id: int) -> bytes:
    row = get_check(check_id)
    if not row:
        return b""
    entities = json.loads(row["entities"] or "[]")
    categories = json.loads(row["categories"] or "{}")
    truths = json.loads(row["facts_true"] or "[]")
    doubts = json.loads(row["facts_doubt"] or "[]")

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    font_ok = False
    for p in [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
    ]:
        try:
            pdfmetrics.registerFont(TTFont("UI", p))
            c.setFont("UI", 12)
            font_ok = True
            break
        except Exception:
            continue
    if not font_ok:
        c.setFont("Helvetica", 12)

    y = 800
    lines = [
        "ChekRez Report",
        f"ID: {row['id']}  Mode: {row['mode']}  Credibility: {row['score']}/100",
        f"Created: {row['created_at']}",
        f"URL: {row['source_url'] or '-'}",
        "Entities: " + (", ".join(entities) if entities else "-"),
        "Categories:",
    ]
    for k, v in categories.items():
        lines.append(f"  - {k}: {v}/100")
    lines.append("Truths:")
    lines.extend([f"  - {x}" for x in truths[:12]])
    lines.append("Doubts:")
    lines.extend([f"  - {x}" for x in doubts[:12]])

    for line in lines:
        if y < 60:
            c.showPage()
            if font_ok:
                c.setFont("UI", 12)
            else:
                c.setFont("Helvetica", 12)
            y = 800
        c.drawString(40, y, line[:115])
        y -= 20

    c.save()
    return buf.getvalue()


@app.route("/")
def index():
    return render_template("index.html", user=current_user())


@app.route("/register", methods=["GET", "POST"])
def register_page():
    warning = ""
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        if len(username) < 3 or "@" not in email or len(password) < 6:
            warning = "Проверь данные: имя от 3 символов, корректный email, пароль от 6 символов."
        else:
            con = sqlite3.connect(DB_PATH)
            try:
                con.execute(
                    "INSERT INTO users (username, email, password_hash, created_at) VALUES (?, ?, ?, ?)",
                    (username, email, generate_password_hash(password), datetime.utcnow().isoformat(timespec="seconds")),
                )
                con.commit()
                uid = con.execute("SELECT id FROM users WHERE email = ?", (email,)).fetchone()[0]
                session["user_id"] = int(uid)
                return redirect(url_for("index"))
            except sqlite3.IntegrityError:
                warning = "Пользователь с таким email или username уже существует."
            finally:
                con.close()
    return render_template("register.html", warning=warning, user=current_user())


@app.route("/login", methods=["GET", "POST"])
def login_page():
    warning = ""
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        con = sqlite3.connect(DB_PATH)
        con.row_factory = sqlite3.Row
        row = con.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
        con.close()
        if row and check_password_hash(row["password_hash"], password):
            session["user_id"] = int(row["id"])
            return redirect(url_for("index"))
        warning = "Неверный email или пароль."
    return render_template("login.html", warning=warning, user=current_user())


@app.route("/logout")
def logout_page():
    session.pop("user_id", None)
    return redirect(url_for("index"))


@app.route("/history")
def history_page():
    if not current_user():
        return redirect(url_for("login_page"))
    rows = get_recent_checks()
    return render_template("history.html", rows=rows, user=current_user())


@app.route("/history/clear", methods=["POST"])
def clear_history_page():
    user = current_user()
    if not user:
        return redirect(url_for("login_page"))
    clear_user_history(int(user["id"]))
    return redirect(url_for("history_page"))


@app.route("/export/<int:check_id>.pdf")
def export_report(check_id: int):
    pdf = render_pdf_report(check_id)
    if not pdf:
        return redirect(url_for("index"))
    return Response(pdf, mimetype="application/pdf", headers={"Content-Disposition": f"inline; filename=checkrez_{check_id}.pdf"})


@app.route("/resume", methods=["GET", "POST"])
def resume_page():
    context = {
        "score": None,
        "truths": [],
        "doubts": [],
        "results": [],
        "warnings": [],
        "entities": [],
        "categories": {},
        "checks": [],
        "explain": [],
        "source_weights": {},
        "check_id": None,
        "company_profile": None,
    }
    if request.method == "POST":
        limit_warn = enforce_rate_limit()
        if limit_warn:
            context["warnings"].append(limit_warn)
            return render_template("resume.html", user=current_user(), **context)

        user = current_user()
        deep = (request.form.get("deep_check") == "on") and bool(user)
        text, warnings, _ = extract_input_text(request.form.get("raw_text", ""), request.form.get("resume_url", ""), request.files.get("resume_file"))
        context["warnings"] = warnings
        if request.form.get("deep_check") == "on" and not user:
            context["warnings"].append("Глубокий поиск доступен только зарегистрированным пользователям.")
        if text:
            score, truths, doubts, results, entities, categories, checks, explain, source_weights = analyze_resume(text, deep=deep)
            check_id = save_check("resume", request.form.get("resume_url", ""), score, entities, categories, truths, doubts)
            context.update(
                {
                    "score": score,
                    "truths": truths,
                    "doubts": doubts,
                    "results": results,
                    "entities": entities,
                    "categories": categories,
                    "checks": checks,
                    "explain": explain,
                    "source_weights": source_weights,
                    "check_id": check_id if check_id else None,
                }
            )
            if not user:
                context["warnings"].append("История не сохраняется без регистрации.")
        else:
            context["warnings"].append("Нужно добавить текст, файл или ссылку для проверки.")
    return render_template("resume.html", user=current_user(), **context)


@app.route("/vacancy", methods=["GET", "POST"])
def vacancy_page():
    context = {
        "score": None,
        "truths": [],
        "doubts": [],
        "results": [],
        "warnings": [],
        "entities": [],
        "categories": {},
        "checks": [],
        "explain": [],
        "source_weights": {},
        "check_id": None,
        "company_profile": None,
    }
    if request.method == "POST":
        limit_warn = enforce_rate_limit()
        if limit_warn:
            context["warnings"].append(limit_warn)
            return render_template("vacancy.html", user=current_user(), **context)

        user = current_user()
        deep = (request.form.get("deep_check") == "on") and bool(user)
        url = request.form.get("vacancy_url", "")
        text, warnings, soup = extract_input_text(request.form.get("raw_text", ""), url, request.files.get("vacancy_file"))
        hh_hints = extract_hh_vacancy_hints(url, soup) if url else VacancyHints()
        context["warnings"] = warnings
        if request.form.get("deep_check") == "on" and not user:
            context["warnings"].append("Глубокий поиск доступен только зарегистрированным пользователям.")
        if text:
            score, truths, doubts, results, entities, categories, checks, explain, source_weights, company_profile = analyze_vacancy(
                text, source_url=url, hh_hints=hh_hints, deep=deep
            )
            check_id = save_check("vacancy", url, score, entities, categories, truths, doubts)
            context.update(
                {
                    "score": score,
                    "truths": truths,
                    "doubts": doubts,
                    "results": results,
                    "entities": entities,
                    "categories": categories,
                    "checks": checks,
                    "explain": explain,
                    "source_weights": source_weights,
                    "check_id": check_id if check_id else None,
                    "company_profile": company_profile,
                }
            )
            if not user:
                context["warnings"].append("История не сохраняется без регистрации.")
        else:
            context["warnings"].append("Нужно добавить текст, файл или ссылку для проверки.")
    return render_template("vacancy.html", user=current_user(), **context)


if __name__ == "__main__":
    init_db()
    app.run(host="127.0.0.1", port=5050, debug=True)
