from __future__ import annotations

import html
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional

_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149 Safari/537.36"
)

_GENERIC_NAMES = {
    "sims", "sims 3", "the sims", "the sims 3", "tumblr", "patreon", "pinterest",
    "download", "downloads", "creator", "converter", "conversion", "mesh", "original",
    "unknown", "package", "custom content", "cc", "ts3", "ts4", "4t3", "2t3", "3t4",
}

# Generic franchise/version words create noisy official-EA and unrelated search hits.
# They are removed from outgoing queries only. Conversion markers such as 4t3/2t3
# are intentionally preserved because they are useful attribution evidence.
_BLOCKED_SEARCH_PATTERNS = (
    r"\bthe\s+sims\s*[234]\b",
    r"\bsims\s*[234]\b",
    r"\bsims[234]\b",
    r"\bts[234]\b",
    r"\bthe\s+sims\b",
    r"\bsims\b",
)

# Words that can help describe a result but are too generic to prove that a web
# result is about the selected package. At least one more distinctive token from
# the user's/package query must appear in a result before CAS Curator shows it.
_SEARCH_IDENTITY_STOPWORDS = {
    "the", "and", "for", "with", "from", "into", "package", "packages",
    "download", "downloads", "custom", "content", "creator", "converter",
    "converted", "conversion", "original", "mesh", "meshes", "made", "version",
    "retexture", "retextured", "by", "cc", "cas", "4t3", "2t3", "3t4", "3t2",
    "hair", "hairs", "clothing", "clothes", "dress", "dresses", "top", "tops",
    "bottom", "bottoms", "shoe", "shoes", "accessory", "accessories", "makeup",
    "male", "female", "adult", "teen", "child", "toddler",
}


def clean_search_query(text: str) -> str:
    """Remove generic Sims/version noise from an outgoing attribution search query."""
    value = str(text or "")
    for pattern in _BLOCKED_SEARCH_PATTERNS:
        value = re.sub(pattern, " ", value, flags=re.IGNORECASE)
    # Common filename/search separators become spaces after blocked terms are removed.
    value = re.sub(r"[_]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip(" \t\r\n-–—_:;,.|/\\")
    return value


def clean_package_terms(filename: str) -> str:
    """Turn a package filename into useful human-readable search terms."""
    raw = Path(str(filename or "").strip()).name
    stem = re.sub(r"\.package$", "", raw, flags=re.IGNORECASE)
    stem = re.sub(r"[\[\]{}()]", " ", stem)
    stem = re.sub(r"[_]+", " ", stem)
    # Keep hyphens inside usernames/names, but space out long separator runs.
    stem = re.sub(r"\s*[-–—]{2,}\s*", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    return clean_search_query(stem)


def _distinctive_query_tokens(query: str) -> List[str]:
    """Return tokens useful for rejecting unrelated search-engine noise."""
    cleaned = clean_search_query(query).casefold()
    tokens = re.findall(r"[a-z0-9@][a-z0-9@.'_-]*", cleaned)
    out: List[str] = []
    seen = set()
    for token in tokens:
        token = token.strip("._-'@")
        if len(token) < 3 or token in _SEARCH_IDENTITY_STOPWORDS or token.isdigit():
            continue
        if token not in seen:
            seen.add(token)
            out.append(token)
    return out


def _result_is_relevant(result: Dict[str, str], query: str) -> bool:
    """Require a distinctive package/query token in each displayed web result."""
    tokens = _distinctive_query_tokens(query)
    if not tokens:
        # With no identifying token, displaying arbitrary search-engine results is
        # more misleading than returning no built-in result. Google browser search
        # remains available for manual investigation.
        return False
    haystack = " ".join(
        str(result.get(key, "") or "") for key in ("title", "snippet", "url")
    ).casefold()
    return any(token in haystack for token in tokens)


def build_search_queries(
    filename: str,
    creator: str = "",
    converter: str = "",
    base_query: str = "",
) -> List[str]:
    clean = clean_package_terms(filename)
    if base_query.strip():
        primary = clean_search_query(base_query)
    else:
        parts = []
        if creator.strip():
            parts.append(clean_search_query(creator.strip()))
        if converter.strip():
            parts.append(clean_search_query(converter.strip()))
        if clean:
            parts.append(clean)
        primary = clean_search_query(" ".join(part for part in parts if part))

    if not primary:
        return []

    # Try the exact cleaned filename/query first. This strongly favours a creator
    # post or download page over broad unrelated search-engine results.
    queries: List[str] = []
    if len(primary.split()) > 1:
        queries.append(f'"{primary}"')
    queries.append(primary)

    # Old CC filenames often contain long IDs/hashes that can make a search too exact.
    # Retry once without those machine-like tokens while preserving creator/item words.
    humanized = re.sub(r"(?i)\b(?:[a-f0-9]{10,}|[0-9]{6,})\b", " ", primary)
    humanized = re.sub(r"\s+", " ", humanized).strip(" -–—_:;,.|/\\")
    if humanized and humanized.casefold() != primary.casefold():
        queries.append(humanized)

    lower = primary.casefold()
    marker = re.search(r"(?i)(?:^|[^a-z0-9])(4t3|2t3|3t4|3t2)(?:[^a-z0-9]|$)", primary)
    if marker:
        queries.append(primary + ' "converted by"')
    elif not any(token in lower for token in ("original by", "mesh by", "original creator")):
        queries.append(primary + ' "mesh by"')

    deduped: List[str] = []
    seen = set()
    for query in queries:
        key = query.casefold().strip()
        if key and key not in seen:
            seen.add(key)
            deduped.append(query)
    return deduped[:3]


def _request_text(url: str, timeout: float = 6.0, data: Optional[bytes] = None) -> str:
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "User-Agent": _USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-GB,en;q=0.9",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        payload = response.read(2_000_000)
        charset = response.headers.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


def _strip_html(text: str) -> str:
    text = re.sub(r"(?is)<script\b.*?</script>", " ", text or "")
    text = re.sub(r"(?is)<style\b.*?</style>", " ", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def _bing_rss_search(query: str, limit: int = 8) -> List[Dict[str, str]]:
    url = "https://www.bing.com/search?format=rss&q=" + urllib.parse.quote_plus(query)
    xml_text = _request_text(url)
    root = ET.fromstring(xml_text)
    results: List[Dict[str, str]] = []
    for item in root.findall(".//item"):
        title = _strip_html(item.findtext("title") or "")
        link = (item.findtext("link") or "").strip()
        snippet = _strip_html(item.findtext("description") or "")
        if title or link:
            results.append({
                "title": title,
                "url": link,
                "snippet": snippet,
                "provider": "Bing",
                "query": query,
            })
        if len(results) >= limit:
            break
    return results


def _unwrap_duckduckgo_url(raw_url: str) -> str:
    """Return the external target from a DuckDuckGo redirect URL when possible."""
    value = html.unescape(str(raw_url or "")).strip()
    if value.startswith("//"):
        value = "https:" + value
    elif value.startswith("/"):
        value = "https://duckduckgo.com" + value
    parsed = urllib.parse.urlparse(value)
    if "duckduckgo.com" in parsed.netloc.casefold() and parsed.query:
        qs = urllib.parse.parse_qs(parsed.query)
        uddg = (qs.get("uddg") or qs.get("u") or [""])[0]
        if uddg:
            return urllib.parse.unquote(uddg)
    return value


def _parse_duckduckgo_html(page: str, query: str, limit: int, provider_label: str) -> List[Dict[str, str]]:
    """Parse DuckDuckGo HTML results. Used by both GET and POST endpoints."""
    links = list(re.finditer(
        r'(?is)<a[^>]+class=["\'][^"\']*result__a[^"\']*["\'][^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
        page,
    ))
    results: List[Dict[str, str]] = []
    for idx, match in enumerate(links):
        raw_url = _unwrap_duckduckgo_url(match.group(1))
        title = _strip_html(match.group(2))
        if not title or not raw_url.lower().startswith(("http://", "https://")):
            continue
        next_start = match.end()
        next_end = links[idx + 1].start() if idx + 1 < len(links) else min(len(page), next_start + 7000)
        block = page[next_start:next_end]
        snippet_match = re.search(
            r'(?is)<(?:a|div)[^>]+class=["\'][^"\']*result__snippet[^"\']*["\'][^>]*>(.*?)</(?:a|div)>',
            block,
        )
        snippet = _strip_html(snippet_match.group(1) if snippet_match else "")
        results.append({
            "title": title,
            "url": raw_url,
            "snippet": snippet,
            "provider": provider_label,
            "query": query,
        })
        if len(results) >= limit:
            break
    return results


def _duckduckgo_html_get_search(query: str, limit: int = 8) -> List[Dict[str, str]]:
    # GET and POST behave differently on some networks, so try both before giving up.
    url = "https://html.duckduckgo.com/html/?kl=uk-en&q=" + urllib.parse.quote_plus(query)
    page = _request_text(url)
    return _parse_duckduckgo_html(page, query, limit, "DuckDuckGo")


def _duckduckgo_html_post_search(query: str, limit: int = 8) -> List[Dict[str, str]]:
    body = urllib.parse.urlencode({"q": query, "kl": "uk-en"}).encode("utf-8")
    page = _request_text("https://html.duckduckgo.com/html/", data=body)
    return _parse_duckduckgo_html(page, query, limit, "DuckDuckGo")


def _duckduckgo_lite_search(query: str, limit: int = 8) -> List[Dict[str, str]]:
    """Fallback to DuckDuckGo Lite, whose markup is simpler than the HTML endpoint."""
    url = "https://lite.duckduckgo.com/lite/?kl=uk-en&q=" + urllib.parse.quote_plus(query)
    page = _request_text(url)
    anchors = list(re.finditer(
        r'(?is)<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
        page,
    ))
    results: List[Dict[str, str]] = []
    for idx, match in enumerate(anchors):
        target = _unwrap_duckduckgo_url(match.group(1))
        title = _strip_html(match.group(2))
        if not title or not target.lower().startswith(("http://", "https://")):
            continue
        parsed = urllib.parse.urlparse(target)
        if "duckduckgo.com" in parsed.netloc.casefold():
            continue
        next_start = match.end()
        next_end = anchors[idx + 1].start() if idx + 1 < len(anchors) else min(len(page), next_start + 3000)
        snippet = _strip_html(page[next_start:next_end])[:700].strip()
        results.append({
            "title": title,
            "url": target,
            "snippet": snippet,
            "provider": "DuckDuckGo Lite",
            "query": query,
        })
        if len(results) >= limit:
            break
    return results


def web_search(query: str, limit: int = 8) -> List[Dict[str, str]]:
    """Best-effort keyless search with several ordinary endpoints and strict relevance filtering."""
    errors: List[str] = []
    providers = (
        ("DuckDuckGo HTML GET", _duckduckgo_html_get_search),
        ("DuckDuckGo Lite", _duckduckgo_lite_search),
        ("DuckDuckGo HTML POST", _duckduckgo_html_post_search),
        ("Bing RSS", _bing_rss_search),
    )
    attempted = 0
    for provider_name, provider in providers:
        attempted += 1
        try:
            raw_results = provider(query, limit=max(limit * 2, 12))
        except Exception as exc:
            errors.append(f"{provider_name}: {exc}")
            continue
        relevant = [result for result in raw_results if _result_is_relevant(result, query)]
        if relevant:
            return relevant[:limit]
    if errors and len(errors) == attempted:
        raise RuntimeError("; ".join(errors))
    return []


def _clean_candidate(raw: str) -> str:
    value = html.unescape(str(raw or ""))
    value = re.sub(r"\s+", " ", value).strip(" \t\r\n:;,.|/\\–—-()[]{}")
    # A broad result snippet can sometimes make a regex capture the next credit
    # phrase itself (for example "conversion — Converted by JaneDoe").
    # Strip that credit prefix rather than presenting it as part of the name.
    value = re.sub(r"(?i)^(?:converted|conversion|created|original|mesh)\s+by\s+", "", value).strip()
    # Stop where prose or a second credit normally begins.
    value = re.split(
        r"(?i)\s+(?:for|from|with|and|but|on|at|via|using|includes?|download|converted|conversion|original|mesh|sims\s*[234]|the\s+sims)\b|(?<=[A-Za-z0-9])[.,]\s+|\s*[|;]\s*|\s+[–—]\s+",
        value,
        maxsplit=1,
    )[0].strip(" :;,.|/\\–—-()[]{}")
    words = value.split()
    if len(words) > 5:
        value = " ".join(words[:5])
    # Leading @ is useful for usernames, trailing punctuation is not.
    if value.casefold() in _GENERIC_NAMES or len(value) < 2 or len(value) > 64:
        return ""
    if not re.search(r"[A-Za-z0-9]", value):
        return ""
    return value


def _extract_explicit_candidates(text: str, url: str, title: str) -> List[Dict[str, object]]:
    found: List[Dict[str, object]] = []
    if not text:
        return found
    patterns = [
        (
            "Converter",
            re.compile(
                r"(?i)\b(?:converted(?:\s+(?:to|for)\s+(?:the\s+)?sims\s*3)?|conversion|(?:4t3|2t3)\s+conversion)\s*(?:by|:|[-–—])\s*(@?[A-Za-z0-9][A-Za-z0-9_'’.\- ]{1,62})"
            ),
            3,
        ),
        (
            "Converter",
            re.compile(
                r"(?i)\b(?:4t3|2t3)\s*(?:by|:|[-–—])\s*(@?[A-Za-z0-9][A-Za-z0-9_'’.\- ]{1,62})"
            ),
            3,
        ),
        (
            "Creator",
            re.compile(
                r"(?i)\b(?:original(?:\s+(?:creator|mesh|cc))?|mesh|created)\s*(?:by|:|[-–—])\s*(@?[A-Za-z0-9][A-Za-z0-9_'’.\- ]{1,62})"
            ),
            3,
        ),
        (
            "Creator",
            re.compile(
                r"(?i)\b(?:original\s+creator|original\s+by)\s*(?:is|:|[-–—])?\s*(@?[A-Za-z0-9][A-Za-z0-9_'’.\- ]{1,62})"
            ),
            3,
        ),
        (
            "Creator",
            re.compile(
                r"(?i)\bcredits?\s+(?:to|:)\s*(@?[A-Za-z0-9][A-Za-z0-9_'’.\- ]{1,62})"
            ),
            2,
        ),
    ]
    for role, pattern, score in patterns:
        for match in pattern.finditer(text):
            name = _clean_candidate(match.group(1))
            if not name:
                continue
            evidence_start = max(0, match.start() - 45)
            evidence_end = min(len(text), match.end() + 65)
            evidence = re.sub(r"\s+", " ", text[evidence_start:evidence_end]).strip()
            found.append({
                "field": role,
                "name": name,
                "score": score,
                "evidence": evidence,
                "url": url,
                "title": title,
                "source_kind": "web",
            })
    return found


def _filename_candidates(filename: str) -> List[Dict[str, object]]:
    stem = re.sub(r"\.package$", "", Path(str(filename or "")).name, flags=re.IGNORECASE)
    if not stem:
        return []
    normalized = re.sub(r"\s+", "_", stem)
    out: List[Dict[str, object]] = []

    # Pattern: OriginalCreator_4t3_Item_by_Converter
    match = re.search(
        r"(?i)(?:^|[_\-])([A-Za-z0-9][A-Za-z0-9.'-]{1,31})[_\-]+(?:4t3|2t3)[_\-]+.*?[_\-]+by[_\-]+([A-Za-z0-9][A-Za-z0-9.'-]{1,31})(?:$|[_\-])",
        normalized,
    )
    if match:
        original = _clean_candidate(match.group(1))
        converter = _clean_candidate(match.group(2))
        if original:
            out.append({"field": "Creator", "name": original, "score": 1, "evidence": f"Filename pattern: {stem}", "url": "", "title": "Package filename", "source_kind": "filename"})
        if converter:
            out.append({"field": "Converter", "name": converter, "score": 1, "evidence": f"Filename pattern: {stem}", "url": "", "title": "Package filename", "source_kind": "filename"})
        return out

    # Pattern: Converter_4t3_OriginalCreator_Item. This is common, but ambiguous,
    # so suggestions from it are always Low confidence and never auto-checked.
    match = re.search(
        r"(?i)^([A-Za-z0-9][A-Za-z0-9.'-]{1,31})[_\-]+(?:4t3|2t3)[_\-]+([A-Za-z0-9][A-Za-z0-9.'-]{1,31})(?:$|[_\-])",
        normalized,
    )
    if match:
        converter = _clean_candidate(match.group(1))
        original = _clean_candidate(match.group(2))
        if converter:
            out.append({"field": "Converter", "name": converter, "score": 1, "evidence": f"Possible filename pattern: {stem}", "url": "", "title": "Package filename", "source_kind": "filename"})
        if original:
            out.append({"field": "Creator", "name": original, "score": 1, "evidence": f"Possible filename pattern: {stem}", "url": "", "title": "Package filename", "source_kind": "filename"})
    return out


def _extract_creator_bylines(result: Dict[str, str]) -> List[Dict[str, object]]:
    """Extract a conservative generic "by NAME" creator credit.

    Generic bylines are weaker than explicit phrases such as "mesh by" or
    "original creator". They only become a visible suggestion later when the
    same name is supported by at least two distinct relevant web results.
    """
    found: List[Dict[str, object]] = []
    title = str(result.get("title", "") or "")
    snippet = str(result.get("snippet", "") or "")
    url = str(result.get("url", "") or "")

    patterns = (
        # Strongest generic case: result title ends in "by NAME".
        (title, re.compile(r"(?i)\bby\s+(@?[A-Za-z0-9][A-Za-z0-9_'’\.\- ]{1,48})\s*(?:\.{3}|…)?\s*$")),
        # Result snippets often contain a short item-title byline followed by
        # Download/Labels/etc. Keep the capture deliberately short.
        (snippet, re.compile(r"(?i)\bby\s+(@?[A-Za-z0-9][A-Za-z0-9_'’\.\- ]{1,48}?)(?=\s+(?:download|downloads|labels?|posted|upload|source|http|www\.|$)|[|;]|\s+[–—]\s+)")),
    )
    for text, pattern in patterns:
        if not text:
            continue
        for match in pattern.finditer(text):
            name = _clean_candidate(match.group(1))
            if not name:
                continue
            evidence_start = max(0, match.start() - 55)
            evidence_end = min(len(text), match.end() + 55)
            evidence = re.sub(r"\s+", " ", text[evidence_start:evidence_end]).strip()
            found.append({
                "field": "Creator",
                "name": name,
                "score": 2,
                "evidence": evidence,
                "url": url,
                "title": title,
                "source_kind": "byline",
            })
    return found


def infer_candidates(filename: str, results: List[Dict[str, str]]) -> List[Dict[str, object]]:
    """Return only conservative web-backed Creator/Converter suggestions.

    Filename patterns are intentionally *not* turned into suggestions. A
    filename can still help find the right pages, but CAS Curator now proposes
    attribution only when the web results contain explicit credit wording or
    repeated matching creator bylines.
    """
    raw_candidates: List[Dict[str, object]] = []
    for result in results:
        title = str(result.get("title", "") or "")
        snippet = str(result.get("snippet", "") or "")
        combined = " — ".join(x for x in (title, snippet) if x)
        raw_candidates.extend(_extract_explicit_candidates(combined, result.get("url", ""), title))
        raw_candidates.extend(_extract_creator_bylines(result))

    aggregated: Dict[tuple, Dict[str, object]] = {}
    for candidate in raw_candidates:
        field = str(candidate.get("field", ""))
        name = str(candidate.get("name", "")).strip()
        if field not in {"Creator", "Converter"} or not name:
            continue
        key = (field, name.casefold())
        bucket = aggregated.get(key)
        if bucket is None:
            bucket = {
                "field": field,
                "name": name,
                "score": 0,
                "best_score": 0,
                "web_sources": set(),
                "explicit_sources": set(),
                "byline_sources": set(),
                "evidence": [],
                "url": "",
                "title": "",
            }
            aggregated[key] = bucket
        score = int(candidate.get("score", 0) or 0)
        bucket["score"] = int(bucket["score"]) + score
        bucket["best_score"] = max(int(bucket["best_score"]), score)
        url = str(candidate.get("url", "") or "")
        source_kind = str(candidate.get("source_kind", "") or "")
        if url:
            bucket["web_sources"].add(url)
            if source_kind == "byline":
                bucket["byline_sources"].add(url)
            elif source_kind == "web":
                bucket["explicit_sources"].add(url)
        if candidate.get("evidence") and str(candidate.get("evidence")) not in bucket["evidence"]:
            bucket["evidence"].append(str(candidate.get("evidence")))
        if not bucket["url"] and url:
            bucket["url"] = url
            bucket["title"] = str(candidate.get("title", ""))

    final: List[Dict[str, object]] = []
    for bucket in aggregated.values():
        field = str(bucket["field"])
        explicit_count = len(bucket["explicit_sources"])
        byline_count = len(bucket["byline_sources"])
        web_source_count = len(bucket["web_sources"])

        # Conservative display threshold:
        # - any explicit attribution phrase is enough to suggest (Medium)
        # - a generic "by NAME" must agree across >=2 distinct relevant results
        # - converter suggestions always require explicit conversion wording
        if field == "Converter":
            if explicit_count < 1:
                continue
        else:
            if explicit_count < 1 and byline_count < 2:
                continue

        if explicit_count >= 2 or (field == "Creator" and byline_count >= 2):
            confidence = "High"
        else:
            confidence = "Medium"

        final.append({
            "field": field,
            "name": bucket["name"],
            "confidence": confidence,
            "evidence": " | ".join(bucket["evidence"][:3]),
            "url": bucket["url"],
            "title": bucket["title"],
            "source_count": web_source_count,
            "explicit_source_count": explicit_count,
            "byline_source_count": byline_count,
            "score": int(bucket["score"]),
        })

    rank = {"High": 0, "Medium": 1}
    final.sort(key=lambda item: (
        item["field"],
        rank.get(str(item["confidence"]), 9),
        -int(item.get("source_count", 0)),
        str(item["name"]).casefold(),
    ))
    return final

def lookup_attribution(
    filename: str,
    creator: str = "",
    converter: str = "",
    base_query: str = "",
) -> Dict[str, object]:
    queries = build_search_queries(filename, creator, converter, base_query)
    results: List[Dict[str, str]] = []
    errors: List[str] = []
    seen_urls = set()
    for query in queries:
        try:
            batch = web_search(query, limit=8)
        except Exception as exc:
            errors.append(f"{query}: {exc}")
            continue
        for result in batch:
            url_key = str(result.get("url", "")).casefold().strip()
            title_key = str(result.get("title", "")).casefold().strip()
            key = url_key or title_key
            if key and key in seen_urls:
                continue
            if key:
                seen_urls.add(key)
            results.append(result)
            if len(results) >= 20:
                break
        if len(results) >= 20:
            break

    candidates = infer_candidates(filename, results)
    return {
        "queries": queries,
        "results": results,
        "candidates": candidates,
        "errors": errors,
    }
