#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
paper_seek_plus / app.py
========================

PaperSeek 改进版：一个 Streamlit 的"自然语言 -> 多库文献检索 -> LLM 重排 -> Excel"
工作台。

设计思想（提炼自 PaperSeek，但不复制其代码）
------------------------------------------------
1. 大模型只做三件事：理解研究问题、生成结构化检索意图、对结果做相关性判断。
2. 真正的检索由程序调用各文献库的 API 或用户自己的机构会话完成，模型不"进库搜索"。
3. 流程由程序控制（确定性骨架），模型是可替换的零件，每个零件都有降级方案。
4. 数据源做成可插拔 connector，单个库失败被隔离，不影响其他库。

合规声明
--------
- 本程序不包含任何真实密钥，所有凭据由用户在界面填写，仅保存在本机 config.json。
- Session/Cookie 模式仅用于用户本人已合法持有的机构订阅访问权限。
- 不实现、不鼓励绕过登录、验证码、付费墙或反爬机制。

运行：streamlit run app.py
"""

from __future__ import annotations

import io
import json
import os
import re
import time
import html as html_lib
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote_plus, urljoin, urlparse

import requests
import pandas as pd
import streamlit as st

try:
    from bs4 import BeautifulSoup  # type: ignore
    BS4_AVAILABLE = True
except Exception:  # pragma: no cover
    BeautifulSoup = None  # type: ignore
    BS4_AVAILABLE = False

try:
    import openpyxl  # noqa: F401
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except Exception:  # pragma: no cover
    OPENPYXL_AVAILABLE = False


# ---------------------------------------------------------------------------
# 全局常量
# ---------------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
HTTP_TIMEOUT = 30
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36 paper-seek-plus/1.0"
)

GRAPH_COLUMNS: List[Tuple[str, str]] = [
    ("文献库", "source"),
    ("标题", "title"),
    ("作者", "authors"),
    ("年份", "year"),
    ("期刊/会议", "venue"),
    ("DOI", "doi"),
    ("链接", "url"),
    ("摘要", "abstract"),
    ("关键词", "keywords"),
    ("相关性分数", "score"),
    ("LLM 摘要/推荐理由", "llm_note"),
    ("检索查询", "query_used"),
    ("数据源类型", "source_type"),
    ("抓取时间", "fetched_at"),
    ("去重标识", "dedup_key"),
    ("原始错误或备注", "note"),
]

RECORD_FIELDS = [
    "source_id", "source", "source_type", "title", "authors", "year", "venue",
    "doi", "url", "abstract", "keywords", "score", "llm_note", "query_used",
    "fetched_at", "dedup_key", "note",
]

HTTP_SESSION = requests.Session()
HTTP_SESSION.headers.update({"User-Agent": USER_AGENT})


# ---------------------------------------------------------------------------
# 异常
# ---------------------------------------------------------------------------

class ConnectorError(Exception):
    """单个文献库层面的可预期错误（配置缺失、鉴权失败、返回不可解析等）。"""


class LLMError(Exception):
    """大模型调用层面的错误。"""


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")
_JSON_BLOCK_RE = re.compile(r"\{.*\}|\[.*\]", re.S)


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_text(value: Any, limit: int = 0) -> str:
    """去掉 HTML 标签、压缩空白，可选截断。"""
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        value = "; ".join(str(v) for v in value if v not in (None, ""))
    text = html_lib.unescape(str(value))
    text = _TAG_RE.sub(" ", text)
    text = _WS_RE.sub(" ", text).strip()
    if limit and len(text) > limit:
        text = text[:limit].rstrip() + "…"
    return text


def to_int(value: Any) -> Optional[int]:
    if value in (None, ""):
        return None
    m = re.search(r"(1[89]\d{2}|20\d{2}|21\d{2})", str(value))
    return int(m.group(1)) if m else None


def normalize_doi(value: Any) -> str:
    text = str(value or "").strip().lower()
    for prefix in ("https://doi.org/", "http://doi.org/", "https://dx.doi.org/", "doi:"):
        if text.startswith(prefix):
            text = text[len(prefix):]
    text = text.strip()
    return text if text.startswith("10.") else ""


def normalize_title(value: Any) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^0-9a-z\u4e00-\u9fff ]+", " ", text)
    return _WS_RE.sub(" ", text).strip()


def mask_secret(value: str, keep: int = 4) -> str:
    """敏感信息掩码显示。"""
    if not value:
        return ""
    value = str(value)
    if len(value) <= keep * 2:
        return "•" * len(value)
    return f"{value[:keep]}{'•' * 8}{value[-keep:]}"


def parse_json_object(text: str) -> Optional[Any]:
    """从模型输出里尽力抽取一个 JSON 对象/数组。"""
    if not text:
        return None
    raw = text.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z]*\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw).strip()
    try:
        return json.loads(raw)
    except Exception:
        pass
    m = _JSON_BLOCK_RE.search(raw)
    if m:
        for candidate in (m.group(0),):
            try:
                return json.loads(candidate)
            except Exception:
                continue
    return None


_PATH_TOKEN_RE = re.compile(r"([^.\[\]]+)|\[(\d+|\*)\]")


def _walk_path(cur: Any, tokens: List[Tuple[str, Any]]) -> List[Any]:
    if not tokens:
        return [cur]
    kind, val = tokens[0]
    rest = tokens[1:]
    out: List[Any] = []
    if kind == "key":
        if isinstance(cur, dict) and val in cur:
            out.extend(_walk_path(cur[val], rest))
    else:
        if isinstance(cur, list):
            if val == "*":
                for item in cur:
                    out.extend(_walk_path(item, rest))
            elif isinstance(val, int) and 0 <= val < len(cur):
                out.extend(_walk_path(cur[val], rest))
    return out


def json_path_get(data: Any, path: str, default: Any = None) -> Any:
    """
    极简 JSON 路径取值，支持：
      $.a.b        -> 取字典链
      $.a[0].b     -> 数组下标
      $.a[*].b     -> 数组全量投影（返回 list）
    多值结果返回 list，单值返回标量。
    """
    if data is None or path in (None, "", "$"):
        return data if path == "$" else default
    p = str(path).strip()
    if p.startswith("$"):
        p = p[1:]
    if p.startswith("."):
        p = p[1:]
    if not p:
        return data
    tokens: List[Tuple[str, Any]] = []
    for m in _PATH_TOKEN_RE.finditer(p):
        if m.group(1) is not None:
            tokens.append(("key", m.group(1)))
        else:
            raw = m.group(2)
            tokens.append(("idx", "*" if raw == "*" else int(raw)))
    values = [v for v in _walk_path(data, tokens) if v not in (None, "")]
    if not values:
        return default
    return values if len(values) > 1 else values[0]


def as_text(value: Any, joiner: str = "; ") -> str:
    if isinstance(value, list):
        parts: List[str] = []
        for item in value:
            if isinstance(item, dict):
                parts.append(str(item.get("name") or item.get("value") or item.get("text") or ""))
            else:
                parts.append(str(item))
        return joiner.join([p for p in parts if p])
    return clean_text(value)


# ---------------------------------------------------------------------------
# 配置读写
# ---------------------------------------------------------------------------

def _source_defaults(preset: Dict[str, Any]) -> Dict[str, Any]:
    """由 preset 生成一个默认源配置。"""
    return {
        "id": preset["id"],
        "display_name": preset["name"],
        "connector": preset.get("connector", "template"),
        "source_type": preset.get("source_type", "api"),
        "preset": True,
        "enabled": bool(preset.get("default_enabled", False)),
        "endpoint": preset.get("endpoint", ""),
        "api_key": "",
        "inst_token": "",
        "auth_mode": preset.get("auth_mode", "header"),
        "key_header": preset.get("key_header", "X-API-Key"),
        "key_param": preset.get("key_param", "apikey"),
        "db": preset.get("db", ""),
        "extra_headers": {},
        "extra_params": {},
        "url_template": preset.get("url_template", ""),
        "method": preset.get("method", "GET"),
        "query_param": preset.get("query_param", "q"),
        "cookie": "",
        "response_type": preset.get("response_type", "json"),
        "field_map": json.loads(json.dumps(preset.get("field_map") or {})),
        "css_map": json.loads(json.dumps(preset.get("css_map") or {})),
        "pages": 1,
        "use_playwright": False,
        "notes": preset.get("notes", ""),
        "verified": bool(preset.get("verified", False)),
        "test": {"ok": None, "msg": "", "ts": ""},
    }


def default_config() -> Dict[str, Any]:
    return {
        "version": 1,
        "llm": {
            "base_url": "https://api.deepseek.com",
            "api_key": "",
            "model": "deepseek-chat",
            "temperature": 0.2,
            "max_tokens": 2048,
            "timeout": 120,
        },
        "sources": {p["id"]: _source_defaults(p) for p in PRESETS},
        "custom_sources": [],
        "settings": {
            "max_results": 30,
            "enable_rerank": True,
            "enable_summary": True,
            "rerank_top_n": 80,
        },
    }


def _merge_source(base: Dict[str, Any], saved: Dict[str, Any]) -> Dict[str, Any]:
    merged = dict(base)
    for key, value in (saved or {}).items():
        if key in ("field_map", "css_map") and isinstance(value, dict):
            merged[key] = value
        else:
            merged[key] = value
    return merged


def load_config() -> Dict[str, Any]:
    cfg = default_config()
    if not CONFIG_PATH.exists():
        return cfg
    try:
        saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return cfg
    if isinstance(saved.get("llm"), dict):
        cfg["llm"].update(saved["llm"])
    if isinstance(saved.get("settings"), dict):
        cfg["settings"].update(saved["settings"])
    saved_sources = saved.get("sources") or {}
    for sid, base in cfg["sources"].items():
        if sid in saved_sources and isinstance(saved_sources[sid], dict):
            cfg["sources"][sid] = _merge_source(base, saved_sources[sid])
    # 保存过、但已不在 preset 列表里的来源，作为自定义来源保留
    known = set(cfg["sources"].keys())
    custom = [c for c in (saved.get("custom_sources") or []) if isinstance(c, dict) and c.get("id")]
    for sid, item in saved_sources.items():
        if sid not in known and isinstance(item, dict):
            item.setdefault("preset", False)
            custom.append(item)
    cfg["custom_sources"] = custom
    return cfg


def save_config(cfg: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        try:
            os.chmod(CONFIG_PATH, 0o600)
        except Exception:
            pass
        return True, f"已保存到 {CONFIG_PATH}"
    except Exception as exc:
        return False, f"保存失败：{exc}"


# ---------------------------------------------------------------------------
# 默认文献库预设
# ---------------------------------------------------------------------------
# verified=True  : 有公开文档的官方 API，字段解析按官方响应结构实现
# verified=False : 无稳定公开 API，或 endpoint 需由机构文档确认，仅提供可配置模板

PRESETS: List[Dict[str, Any]] = [
    {
        "id": "wos",
        "name": "Web of Science",
        "connector": "wos",
        "source_type": "api",
        "endpoint": "https://api.clarivate.com/apis/wos-starter/v1/documents",
        "default_enabled": False,
        "verified": True,
        "db": "WOS",
        "notes": (
            "Clarivate Web of Science Starter API。需机构订阅并在开发者门户申请 API Key，"
            "请求头 X-ApiKey。支持 q / db / limit / page / sortField。"
            "若机构只开通了旧版 Expanded API，请改用 Session 模板。"
        ),
    },
    {
        "id": "scopus",
        "name": "Scopus",
        "connector": "scopus",
        "source_type": "api",
        "endpoint": "https://api.elsevier.com/content/search/scopus",
        "default_enabled": False,
        "verified": True,
        "notes": (
            "Elsevier Scopus Search API。请求头 X-ELS-APIKey；若机构开通了 institutional token，"
            "可另填 X-ELS-Insttoken。返回体为 search-results.entry。"
            "注意：Scopus 搜索接口默认不返回摘要，摘要列可能为空。"
        ),
    },
    {
        "id": "sciencedirect",
        "name": "ScienceDirect",
        "connector": "sciencedirect",
        "source_type": "api",
        "endpoint": "https://api.elsevier.com/content/search/sciencedirect",
        "default_enabled": False,
        "verified": True,
        "notes": (
            "Elsevier ScienceDirect Search API。与 Scopus 同一套 X-ELS-APIKey 鉴权。"
            "该接口需要机构订阅对应的全文/元数据权限，未订阅时通常返回 401/403。"
        ),
    },
    {
        "id": "ieee",
        "name": "IEEE Xplore",
        "connector": "ieee",
        "source_type": "api",
        "endpoint": "https://ieeexploreapi.ieee.org/api/v1/search/articles",
        "default_enabled": False,
        "verified": True,
        "notes": (
            "IEEE Xplore Metadata API。apikey 作为查询参数传入，querytext 支持布尔式，"
            "ranges 参数可限定年份（如 2020_2024_Year），单页最多 200 条。"
        ),
    },
    {
        "id": "springer",
        "name": "SpringerLink",
        "connector": "springer",
        "source_type": "api",
        "endpoint": "https://api.springernature.com/meta/v2/json",
        "default_enabled": False,
        "verified": True,
        "notes": (
            "Springer Nature Metadata API（Meta v2）。api_key 作为查询参数，q 接受布尔式，"
            "p 为每页条数、s 为起始位置。免费额度有限，需在 Springer Nature 开发者门户申请。"
        ),
    },
    {
        "id": "ei",
        "name": "EI Compendex / Engineering Village",
        "connector": "template",
        "source_type": "api",
        "default_enabled": False,
        "verified": False,
        "endpoint": "",
        "response_type": "json",
        "field_map": {
            "container": "$.results[*]",
            "title": "title",
            "authors": "authors[*].name",
            "year": "year",
            "venue": "source",
            "doi": "doi",
            "url": "link",
            "abstract": "abstract",
            "keywords": "keywords[*]",
        },
        "notes": (
            "Engineering Village（Ei Compendex）的检索 API 仅对订阅机构开放，且不同机构的"
            "接入域名/路径不一致，官方文档需从机构图书馆获取。此处不预填 endpoint，"
            "请向图书馆索取后填入，并在下方 field_map 中按实际返回结构调整字段路径。"
            "若机构只提供网页检索，请把本库切到 Session 模式（下方 url_template + CSS 映射）。"
        ),
    },
    {
        "id": "informs",
        "name": "INFORMS Online Journals",
        "connector": "template",
        "source_type": "session",
        "default_enabled": False,
        "verified": False,
        "url_template": "https://pubsonline.informs.org/action/doSearch?AllField={query}",
        "response_type": "html",
        "css_map": {
            "container": "div.issue-item, li.search__item, article",
            "title": "h2 a, h3 a, .issue-item__title a",
            "authors": ".issue-item__authors, .search__item-authors, .entryAuthor",
            "link": "h2 a, h3 a, a.issue-item__title",
            "abstract": ".issue-item__abstract, .abstractSection",
        },
        "notes": (
            "INFORMS 期刊平台（Atypon Literatum）没有公开检索 API。默认走 Session 模式："
            "在浏览器登录机构账号后复制 Cookie 填入。站点为服务端渲染，requests 通常可直接取到列表；"
            "若返回 0 条或登录页，说明 Cookie 失效或页面改版，需要在下方调整 CSS 选择器。"
        ),
    },
    {
        "id": "acm",
        "name": "ACM Digital Library",
        "connector": "template",
        "source_type": "session",
        "default_enabled": False,
        "verified": False,
        "url_template": "https://dl.acm.org/action/doSearch?AllField={query}",
        "response_type": "html",
        "css_map": {
            "container": "div.issue-item, li.search__item, article",
            "title": "h2 a, h3 a, .issue-item__title a",
            "authors": ".issue-item__authors, .search__item-authors, .entryAuthor",
            "link": "h2 a, h3 a, a.issue-item__title",
            "abstract": ".issue-item__abstract, .abstractSection",
        },
        "notes": (
            "ACM DL 没有公开的检索 API（仅对签约机构提供 TDM 文本挖掘接口，需单独申请）。"
            "默认 Session 模式，同样基于 Atypon Literatum。ACM 对自动化访问较敏感，"
            "请控制请求频率；如遇拦截请改用官方 TDM 接口或手工导出。"
        ),
    },
    {
        "id": "cnki",
        "name": "CNKI 中国知网",
        "connector": "template",
        "source_type": "session",
        "default_enabled": False,
        "verified": False,
        "url_template": "https://kns.cnki.net/kns8s/defaultresult/index?kw={query}",
        "response_type": "html",
        "css_map": {
            "container": "table.result-table-list tbody tr, .result-table-list tr",
            "title": "td.name a, .name a",
            "authors": "td.author, .author",
            "link": "td.name a, .name a",
        },
        "notes": (
            "CNKI 没有公开检索 API。默认 Session 模式。CNKI 结果页大量依赖 JavaScript 渲染，"
            "requests 往往只能拿到空壳 HTML；若下方选择 Playwright 渲染仍取不到数据，"
            "说明站点结构已变或需要验证码，请改用手工导出后合并。"
        ),
    },
    # ---- 公共开放源：无需任何密钥，用于开箱验证整条流水线 ----
    {
        "id": "openalex",
        "name": "OpenAlex（开放源，免密钥）",
        "connector": "openalex",
        "source_type": "api",
        "endpoint": "https://api.openalex.org/works",
        "default_enabled": True,
        "verified": True,
        "notes": "完全开放，无需密钥。可选填 email 进入 polite pool。支持年份过滤与摘要重建。",
    },
    {
        "id": "crossref",
        "name": "Crossref（开放源，免密钥）",
        "connector": "crossref",
        "source_type": "api",
        "endpoint": "https://api.crossref.org/works",
        "default_enabled": True,
        "verified": True,
        "notes": "完全开放，无需密钥。DOI 注册库，元数据齐全，排序能力弱于专业索引库。",
    },
]

PRESET_BY_ID = {p["id"]: p for p in PRESETS}


# ---------------------------------------------------------------------------
# HTTP 基础
# ---------------------------------------------------------------------------

def http_request(
    method: str,
    url: str,
    *,
    params: Optional[Dict[str, Any]] = None,
    headers: Optional[Dict[str, str]] = None,
    json_body: Optional[Any] = None,
    data: Optional[Any] = None,
    timeout: int = HTTP_TIMEOUT,
    retries: int = 1,
) -> requests.Response:
    last_exc: Optional[Exception] = None
    for attempt in range(retries + 1):
        try:
            resp = HTTP_SESSION.request(
                method.upper(),
                url,
                params=params,
                headers=headers,
                json=json_body,
                data=data,
                timeout=timeout,
                allow_redirects=True,
            )
            if resp.status_code in (429, 500, 502, 503, 504) and attempt < retries:
                time.sleep(1.2 * (attempt + 1))
                continue
            return resp
        except requests.RequestException as exc:
            last_exc = exc
            if attempt < retries:
                time.sleep(1.2 * (attempt + 1))
                continue
            raise ConnectorError(f"网络请求失败：{exc.__class__.__name__}: {exc}") from exc
    raise ConnectorError(f"网络请求失败：{last_exc}")


def describe_http_error(resp: requests.Response, source_name: str) -> str:
    code = resp.status_code
    hint = ""
    if code in (401, 403):
        hint = "（鉴权失败：API Key 无效、过期，或机构未订阅该库）"
    elif code == 429:
        hint = "（触发限流，请降低频率或稍后重试）"
    elif code >= 500:
        hint = "（服务端错误，可稍后重试）"
    body = clean_text(resp.text, 200)
    return f"{source_name} HTTP {code}{hint}：{body}"


# ---------------------------------------------------------------------------
# 大模型客户端（OpenAI 兼容）
# ---------------------------------------------------------------------------

class LLMClient:
    """
    只依赖 /chat/completions 这一最通用的协议，可对接 DeepSeek / Qwen / Moonshot /
    Ollama / vLLM / OpenRouter / 任何 OpenAI 兼容网关。
    """

    def __init__(self, cfg: Dict[str, Any]):
        self.base_url = str(cfg.get("base_url") or "").strip()
        self.api_key = str(cfg.get("api_key") or "").strip()
        self.model = str(cfg.get("model") or "").strip()
        self.temperature = float(cfg.get("temperature", 0.2) or 0.2)
        self.max_tokens = int(cfg.get("max_tokens", 2048) or 2048)
        self.timeout = int(cfg.get("timeout", 120) or 120)

    # -- 端点兼容 ---------------------------------------------------------
    def _candidate_urls(self) -> List[str]:
        base = self.base_url.rstrip("/")
        if not base:
            return []
        if base.endswith("/chat/completions"):
            return [base]
        urls = [f"{base}/chat/completions"]
        if not base.endswith("/v1"):
            urls.append(f"{base}/v1/chat/completions")
        return urls

    def _headers(self) -> Dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        return headers

    def is_configured(self) -> bool:
        return bool(self.base_url and self.model)

    def chat(self, messages: List[Dict[str, str]], temperature: Optional[float] = None,
             max_tokens: Optional[int] = None) -> str:
        if not self.is_configured():
            raise LLMError("AI API 未配置：请先在「AI API 配置」页填写 base_url 与 model。")
        body = {
            "model": self.model,
            "messages": messages,
            "temperature": self.temperature if temperature is None else temperature,
            "max_tokens": max_tokens or self.max_tokens,
        }
        errors: List[str] = []
        for url in self._candidate_urls():
            try:
                resp = http_request(
                    "POST", url, headers=self._headers(), json_body=body,
                    timeout=self.timeout, retries=0,
                )
            except ConnectorError as exc:
                errors.append(str(exc))
                continue
            if resp.status_code >= 300:
                errors.append(f"HTTP {resp.status_code}: {clean_text(resp.text, 200)}")
                # 404 说明路径不对，尝试下一个候选
                if resp.status_code == 404:
                    continue
                break
            try:
                payload = resp.json()
            except Exception:
                errors.append("响应不是合法 JSON")
                continue
            text = self._extract_text(payload)
            if text:
                return text
            errors.append("响应中没有 message.content")
        raise LLMError("大模型调用失败：" + " | ".join(errors[-3:]))

    @staticmethod
    def _extract_text(payload: Dict[str, Any]) -> str:
        choices = payload.get("choices") or []
        if choices:
            msg = choices[0].get("message") or {}
            content = msg.get("content")
            if isinstance(content, str) and content.strip():
                return content
            if isinstance(content, list):
                parts = [c.get("text", "") for c in content if isinstance(c, dict)]
                joined = "".join(parts).strip()
                if joined:
                    return joined
            reasoning = msg.get("reasoning_content")
            if isinstance(reasoning, str) and reasoning.strip():
                return reasoning
            delta = (choices[0].get("delta") or {}).get("content")
            if isinstance(delta, str) and delta.strip():
                return delta
        if isinstance(payload.get("output_text"), str):
            return payload["output_text"]
        return ""

    def test_connection(self) -> Tuple[bool, str]:
        try:
            reply = self.chat(
                [
                    {"role": "system", "content": "你是一个连通性测试助手，请只回复一句很短的话。"},
                    {"role": "user", "content": "请回复：连接成功"},
                ],
                temperature=0.0,
                max_tokens=32,
            )
        except LLMError as exc:
            return False, str(exc)
        return True, reply.strip()[:300]


# ---------------------------------------------------------------------------
# 检索计划（QueryPlan）
# ---------------------------------------------------------------------------

PLAN_SYSTEM_PROMPT = """你是一位资深的学术文献检索策略师。把用户的研究方向描述转换成严格 JSON 检索计划。

只输出 JSON 对象本身，不要 markdown 代码块，不要任何解释文字。JSON 结构必须是：

{
  "topic": "一句话概括研究主题（中文）",
  "keywords_zh": ["中文核心关键词", "..."],
  "keywords_en": ["english core keyword", "..."],
  "synonyms": ["同义词、近义词、缩写、相邻术语", "..."],
  "boolean_queries": {
    "default": "",
    "wos": "",
    "scopus": "",
    "ieee": "",
    "sciencedirect": "",
    "springer": "",
    "cnki": ""
  },
  "year_range": {"start": null, "end": null},
  "include_terms": ["必须包含的概念"],
  "exclude_terms": ["应排除的概念"]
}

各库查询式写法要求：
- default：英文关键词用带引号的短语，同义词用 OR 并加括号，不同概念用 AND。
  例：("large language model" OR "LLM") AND ("literature review" OR "systematic review")
- wos：Web of Science 语法，用 TS=(...) 组合，禁用 NEAR/SAME 与 WC=，例：TS=("large language model" OR LLM) AND TS=(review)
- scopus：用 TITLE-ABS-KEY(...) 组合，例：TITLE-ABS-KEY("large language model" OR "LLM") AND TITLE-ABS-KEY(review)
- ieee：纯布尔式，短语加双引号，例：("large language model" OR "LLM") AND ("review")
- sciencedirect：Elsevier 自由文本语法，例：("large language model" OR "LLM") AND (review)
- springer：自然语言关键词组合，短语加双引号
- cnki：中文关键词，多个词之间用空格，例：大语言模型 文献综述

规则：
1. 一律输出英文学术关键词；中文关键词只用于 CNKI 与界面展示。
2. 不要编造年份。用户没有明确说时间范围时，year_range 的 start/end 都填 null。
3. core 概念 2-5 个，同义词不超过 10 个。
4. 所有 boolean_queries 字段都必须给出非空字符串；如果确实不适用，就用 default 的内容。
"""


@dataclass
class QueryPlan:
    topic: str = ""
    keywords_zh: List[str] = field(default_factory=list)
    keywords_en: List[str] = field(default_factory=list)
    synonyms: List[str] = field(default_factory=list)
    boolean_queries: Dict[str, str] = field(default_factory=dict)
    year_start: Optional[int] = None
    year_end: Optional[int] = None
    include_terms: List[str] = field(default_factory=list)
    exclude_terms: List[str] = field(default_factory=list)
    degraded: bool = False
    raw: Dict[str, Any] = field(default_factory=dict)

    def query_for(self, source_id: str) -> str:
        q = (self.boolean_queries or {}).get(source_id)
        if q and q.strip():
            return q.strip()
        q = (self.boolean_queries or {}).get("default")
        if q and q.strip():
            return q.strip()
        terms = self.keywords_en or self.keywords_zh
        if not terms:
            return self.topic
        return " AND ".join(f'"{t}"' if " " in t else t for t in terms[:4])


def _as_str_list(value: Any, limit: int = 12) -> List[str]:
    if isinstance(value, str):
        parts = re.split(r"[;,，、\n]", value)
        return [p.strip() for p in parts if p.strip()][:limit]
    if isinstance(value, (list, tuple)):
        out = []
        for item in value:
            text = clean_text(item)
            if text:
                out.append(text)
            if len(out) >= limit:
                break
        return out
    return []


def plan_from_dict(data: Dict[str, Any]) -> QueryPlan:
    bq_raw = data.get("boolean_queries")
    bq: Dict[str, str] = {}
    if isinstance(bq_raw, dict):
        for key, value in bq_raw.items():
            if isinstance(value, str) and value.strip():
                bq[str(key).strip().lower()] = value.strip()
    elif isinstance(bq_raw, str) and bq_raw.strip():
        bq["default"] = bq_raw.strip()

    year_raw = data.get("year_range")
    year_start = year_end = None
    if isinstance(year_raw, dict):
        year_start = to_int(year_raw.get("start"))
        year_end = to_int(year_raw.get("end"))
    elif isinstance(year_raw, (list, tuple)) and len(year_raw) == 2:
        year_start, year_end = to_int(year_raw[0]), to_int(year_raw[1])

    kw_en = _as_str_list(data.get("keywords_en"))
    kw_zh = _as_str_list(data.get("keywords_zh"))
    topic = clean_text(data.get("topic"))

    if not bq.get("default"):
        terms = kw_en or kw_zh
        if terms:
            bq["default"] = " AND ".join(f'"{t}"' if " " in t else t for t in terms[:4])
        elif topic:
            bq["default"] = topic

    return QueryPlan(
        topic=topic or (kw_en[0] if kw_en else ""),
        keywords_zh=kw_zh,
        keywords_en=kw_en,
        synonyms=_as_str_list(data.get("synonyms")),
        boolean_queries=bq,
        year_start=year_start,
        year_end=year_end,
        include_terms=_as_str_list(data.get("include_terms")),
        exclude_terms=_as_str_list(data.get("exclude_terms")),
        degraded=False,
        raw=data,
    )


_STOPWORDS_EN = {
    "a", "an", "the", "of", "for", "in", "on", "to", "and", "or", "with", "by",
    "from", "at", "as", "is", "are", "be", "study", "research", "paper", "based",
    "using", "use", "approach", "method", "analysis",
}
_STOPWORDS_ZH = {
    "研究", "分析", "方法", "基于", "探讨", "问题", "影响", "作用", "应用", "的", "与", "和",
    "以及", "进行", "相关", "方面", "情况", "内容", "主要", "一些", "这个", "那个",
}

# 中文降级分词用的连接词/虚词：先按这些词切断，再按标点切断，
# 避免出现「大语言模型在管理」这种跨词边界的假关键词。
_ZH_CONNECTORS = [
    "特别是", "尤其是", "尤其是关于", "以及", "还有", "或者", "并且", "同时",
    "关于", "对于", "基于", "通过", "结合", "应用在", "应用于", "用于", "面向",
    "我想了解", "我想知道", "希望了解", "重点关注", "主要关注", "关注",
    "在", "的", "与", "和", "及", "中", "上", "下", "对", "为", "把", "被", "从", "到",
    "等", "是", "有", "了", "着", "过", "也", "都", "很", "更", "最",
    "以后", "之前", "以来", "以内", "年前", "年后", "近年", "近几年", "当前",
]


def _split_zh_terms(text: str) -> List[str]:
    """无分词库时的中文关键词切分：先剔除连接词与标点，再取 2-8 字的片段。"""
    working = text
    for word in sorted(_ZH_CONNECTORS, key=len, reverse=True):
        working = working.replace(word, " ")
    fragments = re.split(r"[^\u4e00-\u9fffA-Za-z0-9]+", working)
    out: List[str] = []
    for frag in fragments:
        frag = frag.strip()
        if not (2 <= len(frag) <= 10):
            continue
        if frag in _STOPWORDS_ZH or re.fullmatch(r"\d+", frag):
            continue
        out.append(frag)
    return out


def rule_based_plan(user_text: str, year_start: Optional[int] = None,
                    year_end: Optional[int] = None) -> QueryPlan:
    """LLM 不可用时的降级方案：纯规则抽词。"""
    text = clean_text(user_text)
    zh_terms = _split_zh_terms(text)
    if not zh_terms:
        zh_terms = [t for t in re.findall(r"[\u4e00-\u9fff]{2,8}", text) if t not in _STOPWORDS_ZH]
    en_terms = [t.lower() for t in re.findall(r"[A-Za-z][A-Za-z\-]{2,}", text)
                if t.lower() not in _STOPWORDS_EN]

    def dedup(seq: List[str]) -> List[str]:
        seen, out = set(), []
        for item in seq:
            if item and item not in seen:
                seen.add(item)
                out.append(item)
        return out

    zh_terms = dedup(zh_terms)[:8]
    en_terms = dedup(en_terms)[:8]

    default_terms = en_terms[:4] if en_terms else zh_terms[:4]
    default_q = " AND ".join(f'"{t}"' if " " in t else t for t in default_terms) if default_terms else text
    cnki_q = " ".join(zh_terms[:4]) if zh_terms else text

    bq = {
        "default": default_q,
        "wos": " AND ".join(f'TS=("{t}")' if " " in t else f"TS=({t})" for t in default_terms) or default_q,
        "scopus": " AND ".join(f'TITLE-ABS-KEY("{t}")' if " " in t else f"TITLE-ABS-KEY({t})" for t in default_terms) or default_q,
        "ieee": default_q,
        "sciencedirect": default_q,
        "springer": " ".join(f'"{t}"' if " " in t else t for t in default_terms) or default_q,
        "cnki": cnki_q,
    }
    return QueryPlan(
        topic=text[:120],
        keywords_zh=zh_terms,
        keywords_en=en_terms,
        synonyms=[],
        boolean_queries={k: v for k, v in bq.items() if v},
        year_start=year_start,
        year_end=year_end,
        degraded=True,
        raw={},
    )


def build_query_plan(user_text: str, llm: Optional[LLMClient],
                     log: Callable[[str], None] = lambda m: None) -> QueryPlan:
    """调用 LLM 生成检索计划，失败重试一次，再失败降级为规则抽取。"""
    if llm is None or not llm.is_configured():
        log("AI API 未配置，使用规则法生成关键词。")
        return rule_based_plan(user_text)

    messages = [
        {"role": "system", "content": PLAN_SYSTEM_PROMPT},
        {"role": "user", "content": f"研究方向描述：\n{user_text}\n\n请输出 JSON 检索计划。"},
    ]
    for attempt in (1, 2):
        try:
            log(f"正在调用大模型解析检索意图（第 {attempt} 次）…")
            raw = llm.chat(messages, temperature=0.1, max_tokens=1200)
        except LLMError as exc:
            log(f"意图解析调用失败：{exc}")
            break
        data = parse_json_object(raw)
        if isinstance(data, dict):
            plan = plan_from_dict(data)
            if plan.boolean_queries:
                log("检索意图解析完成。")
                return plan
            log("模型返回的 JSON 缺少 boolean_queries，准备重试。")
        else:
            log("模型输出不是合法 JSON，准备重试。")
        messages = messages + [
            {"role": "assistant", "content": raw[:2000]},
            {"role": "user", "content": "上面的输出不是合法 JSON。请只输出一个 JSON 对象，不要代码块、不要解释。"},
        ]

    log("意图解析降级为规则关键词抽取。")
    return rule_based_plan(user_text)


# ---------------------------------------------------------------------------
# 结果记录
# ---------------------------------------------------------------------------

def make_record(**kwargs: Any) -> Dict[str, Any]:
    rec = {key: "" for key in RECORD_FIELDS}
    rec["fetched_at"] = now_str()
    for key, value in kwargs.items():
        if key in rec:
            rec[key] = value
        elif key == "summary_reason":
            rec["llm_note"] = value
    rec["title"] = clean_text(rec.get("title"), 500)
    rec["authors"] = clean_text(rec.get("authors"), 600)
    rec["venue"] = clean_text(rec.get("venue"), 300)
    rec["abstract"] = clean_text(rec.get("abstract"), 4000)
    rec["keywords"] = clean_text(rec.get("keywords"), 400)
    rec["doi"] = normalize_doi(rec.get("doi"))
    rec["url"] = str(rec.get("url") or "").strip()
    year = to_int(rec.get("year"))
    rec["year"] = year if year else ""
    rec["source_type"] = rec.get("source_type") or "api"
    rec["dedup_key"] = compute_dedup_key(rec)
    return rec


def compute_dedup_key(rec: Dict[str, Any]) -> str:
    doi = normalize_doi(rec.get("doi"))
    if doi:
        return f"doi:{doi}"
    title = normalize_title(rec.get("title"))
    if title:
        return f"title:{title[:120]}"
    url = str(rec.get("url") or "").strip().lower()
    if url:
        return f"url:{url[:160]}"
    return ""


# ---------------------------------------------------------------------------
# Connector 基类
# ---------------------------------------------------------------------------

class BaseConnector:
    """统一文献库接口。子类只需实现 search()，必要时覆盖 test_connection()。"""

    connector_type = "base"
    # 该库的检索接口是否真的理解布尔表达式。
    # Crossref 等只做词袋匹配，传 "A AND B" 会把 AND 也当成检索词，因此置 False。
    supports_boolean = True

    def __init__(self, cfg: Dict[str, Any]):
        self.cfg = cfg or {}

    # -- 元信息 -----------------------------------------------------------
    @property
    def id(self) -> str:
        return str(self.cfg.get("id") or "")

    @property
    def display_name(self) -> str:
        return str(self.cfg.get("display_name") or self.id)

    @property
    def source_type(self) -> str:
        return str(self.cfg.get("source_type") or "api")

    # -- 配置 -------------------------------------------------------------
    def extra_headers(self) -> Dict[str, str]:
        raw = self.cfg.get("extra_headers") or {}
        return {str(k): str(v) for k, v in raw.items()} if isinstance(raw, dict) else {}

    def extra_params(self) -> Dict[str, Any]:
        raw = self.cfg.get("extra_params") or {}
        return dict(raw) if isinstance(raw, dict) else {}

    def require(self, field_name: str, label: str) -> str:
        value = str(self.cfg.get(field_name) or "").strip()
        if not value:
            raise ConnectorError(f"{self.display_name} 未配置「{label}」，请先在「文献库配置」页填写。")
        return value

    # -- 生命周期 ---------------------------------------------------------
    def test_connection(self) -> Tuple[bool, str]:
        try:
            recs = self.search("machine learning", max_results=1, year_range=None)
        except ConnectorError as exc:
            return False, str(exc)
        except Exception as exc:  # noqa: BLE001
            return False, f"{exc.__class__.__name__}: {exc}"
        if recs:
            return True, f"连接正常，测试检索返回 {len(recs)} 条记录。"
        return True, "接口可访问，但测试检索没有返回记录（可能是查询词无命中）。"

    def search(self, query: str, max_results: int = 20,
               year_range: Optional[Tuple[Optional[int], Optional[int]]] = None) -> List[Dict[str, Any]]:
        raise NotImplementedError


# ---------------------------------------------------------------------------
# 具体 Connector 实现
# ---------------------------------------------------------------------------

class OpenAlexConnector(BaseConnector):
    connector_type = "openalex"

    @staticmethod
    def _rebuild_abstract(inverted: Any) -> str:
        if not isinstance(inverted, dict) or not inverted:
            return ""
        positions: List[Tuple[int, str]] = []
        for word, idxs in inverted.items():
            if isinstance(idxs, list):
                for idx in idxs:
                    if isinstance(idx, int):
                        positions.append((idx, word))
        positions.sort()
        return clean_text(" ".join(w for _, w in positions), 4000)

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or "https://api.openalex.org/works").strip()
        filters: List[str] = ["type:article|preprint|review"]
        if year_range and any(year_range):
            start, end = year_range
            if start:
                filters.append(f"from_publication_date:{start}-01-01")
            if end:
                filters.append(f"to_publication_date:{end}-12-31")
        params: Dict[str, Any] = {
            "search": q,
            "per-page": min(max(1, max_results), 200),
            "filter": ",".join(filters),
            "sort": "relevance_score:desc",
        }
        email = str(self.cfg.get("inst_token") or "").strip()
        if email:
            params["mailto"] = email
        params.update(self.extra_params())
        resp = http_request("GET", endpoint, params=params, headers=self.extra_headers())
        if resp.status_code != 200:
            raise ConnectorError(describe_http_error(resp, self.display_name))
        try:
            data = resp.json()
        except Exception as exc:
            raise ConnectorError(f"返回不是 JSON：{exc}") from exc

        records: List[Dict[str, Any]] = []
        for item in (data.get("results") or [])[:max_results]:
            authors = [
                ((a.get("author") or {}).get("display_name") or "")
                for a in (item.get("authorships") or [])
            ]
            venue = ""
            loc = item.get("primary_location") or {}
            if isinstance(loc, dict):
                venue = ((loc.get("source") or {}).get("display_name") or "")
            keywords = [k.get("display_name") for k in (item.get("keywords") or []) if isinstance(k, dict)]
            records.append(make_record(
                source_id=self.id,
                source=self.display_name,
                source_type="api",
                title=item.get("title") or item.get("display_name"),
                authors="; ".join([a for a in authors if a][:20]),
                year=item.get("publication_year"),
                venue=venue,
                doi=item.get("doi"),
                url=(item.get("doi") or item.get("id") or ""),
                abstract=self._rebuild_abstract(item.get("abstract_inverted_index")),
                keywords="; ".join([k for k in keywords if k]),
                score="",
                query_used=q,
                note=f"引用次数 {item.get('cited_by_count', '')}",
            ))
        return records


class CrossrefConnector(BaseConnector):
    connector_type = "crossref"
    supports_boolean = False   # Crossref 做词袋匹配，不解析 AND/OR 算子

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or "https://api.crossref.org/works").strip()
        params: Dict[str, Any] = {
            "query.bibliographic": q,
            "rows": min(max(1, max_results), 100),
            "select": "DOI,title,author,issued,container-title,URL,abstract,subject,type,is-referenced-by-count",
        }
        filters = [
            # 白名单限定文献类型：否则 query.bibliographic 会把 "Review for ..." 这类
            # 同行评审记录也检索出来（实测复现），污染结果集。
            "type:journal-article",
            "type:proceedings-article",
            "type:posted-content",
        ]
        if year_range and any(year_range):
            start, end = year_range
            if start:
                filters.append(f"from-pub-date:{start}-01-01")
            if end:
                filters.append(f"until-pub-date:{end}-12-31")
        params["filter"] = ",".join(filters)
        email = str(self.cfg.get("inst_token") or "").strip()
        if email:
            params["mailto"] = email
        params.update(self.extra_params())
        resp = http_request("GET", endpoint, params=params, headers=self.extra_headers())
        if resp.status_code != 200:
            raise ConnectorError(describe_http_error(resp, self.display_name))
        try:
            data = resp.json()
        except Exception as exc:
            raise ConnectorError(f"返回不是 JSON：{exc}") from exc

        items = ((data.get("message") or {}).get("items") or [])[:max_results]
        records: List[Dict[str, Any]] = []
        for item in items:
            title = item.get("title") or []
            container = item.get("container-title") or []
            authors = []
            for author in (item.get("author") or [])[:20]:
                name = " ".join([x for x in [author.get("given"), author.get("family")] if x]).strip()
                if not name:
                    name = author.get("name") or ""
                if name:
                    authors.append(name)
            issued = item.get("issued") or {}
            parts = issued.get("date-parts") or [[None]]
            year = parts[0][0] if parts and parts[0] else None
            records.append(make_record(
                source_id=self.id,
                source=self.display_name,
                source_type="api",
                title=(title[0] if title else ""),
                authors="; ".join(authors),
                year=year,
                venue=(container[0] if container else ""),
                doi=item.get("DOI"),
                url=item.get("URL"),
                abstract=item.get("abstract"),
                keywords="; ".join(item.get("subject") or []),
                query_used=q,
                note=f"被引 {item.get('is-referenced-by-count', '')}",
            ))
        return records


class WosConnector(BaseConnector):
    """Clarivate Web of Science Starter API。"""

    connector_type = "wos"

    def _headers(self) -> Dict[str, str]:
        key = self.require("api_key", "API Key")
        headers = {"X-ApiKey": key, "Accept": "application/json"}
        headers.update(self.extra_headers())
        return headers

    @staticmethod
    def _apply_year(query: str, year_range: Optional[Tuple[Optional[int], Optional[int]]]) -> str:
        if not year_range or not any(year_range):
            return query
        start, end = year_range
        if start and end:
            return f"({query}) AND PY={start}-{end}"
        if start:
            return f"({query}) AND PY={start}-{datetime.now().year}"
        if end:
            return f"({query}) AND PY=1900-{end}"
        return query

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or "https://api.clarivate.com/apis/wos-starter/v1/documents").strip()
        q = self._apply_year(q, year_range)
        records: List[Dict[str, Any]] = []
        page = 1
        limit = 50
        while len(records) < max_results and page <= 4:
            params = {
                "q": q,
                "db": str(self.cfg.get("db") or "WOS"),
                "limit": min(limit, max_results - len(records)),
                "page": page,
            }
            params.update(self.extra_params())
            resp = http_request("GET", endpoint, params=params, headers=self._headers())
            if resp.status_code != 200:
                raise ConnectorError(describe_http_error(resp, self.display_name))
            try:
                data = resp.json()
            except Exception as exc:
                raise ConnectorError(f"返回不是 JSON：{exc}") from exc

            hits = data.get("hits") or data.get("records") or []
            if not hits:
                break
            for item in hits:
                source = item.get("source") or {}
                names = item.get("names") or {}
                identifiers = item.get("identifiers") or {}
                keywords = item.get("keywords") or {}
                authors = [
                    a.get("displayName") or a.get("wosStandard") or ""
                    for a in (names.get("authors") or [])
                    if isinstance(a, dict)
                ]
                links = item.get("links") or {}
                url = ""
                if isinstance(links, dict):
                    url = links.get("record") or links.get("references") or ""
                records.append(make_record(
                    source_id=self.id,
                    source=self.display_name,
                    source_type="api",
                    title=item.get("title"),
                    authors="; ".join([a for a in authors if a][:20]),
                    year=source.get("publishYear"),
                    venue=source.get("sourceTitle") or "",
                    doi=identifiers.get("doi"),
                    url=url,
                    abstract="",
                    keywords="; ".join(list(keywords.get("authorKeywords") or [])[:10]),
                    query_used=q,
                    note=f"WoS 记录号 {item.get('uid', '')}",
                ))
                if len(records) >= max_results:
                    break
            total = (data.get("metadata") or {}).get("total")
            if isinstance(total, int) and len(records) >= total:
                break
            page += 1
        return records


class _ElsevierBase(BaseConnector):
    """Elsevier 系（Scopus / ScienceDirect）共用鉴权与分页逻辑。"""

    default_endpoint = ""
    max_per_page = 25

    def _headers(self) -> Dict[str, str]:
        key = self.require("api_key", "API Key")
        headers = {"X-ELS-APIKey": key, "Accept": "application/json"}
        token = str(self.cfg.get("inst_token") or "").strip()
        if token:
            headers["X-ELS-Insttoken"] = token
        headers.update(self.extra_headers())
        return headers

    @staticmethod
    def _apply_year(query: str, year_range: Optional[Tuple[Optional[int], Optional[int]]]) -> str:
        if not year_range or not any(year_range):
            return query
        start, end = year_range
        parts = [query]
        if start:
            parts.append(f"PUBYEAR > {start - 1}")
        if end:
            parts.append(f"PUBYEAR < {end + 1}")
        return " AND ".join(f"({p})" if " " in p else p for p in parts)


class ScopusConnector(_ElsevierBase):
    connector_type = "scopus"
    default_endpoint = "https://api.elsevier.com/content/search/scopus"

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or self.default_endpoint).strip()
        q = self._apply_year(q, year_range)
        records: List[Dict[str, Any]] = []
        start = 0
        while len(records) < max_results and start <= 200:
            params = {
                "query": q,
                "count": min(self.max_per_page, max_results - len(records)),
                "start": start,
                "sort": "-coverDate",
            }
            params.update(self.extra_params())
            resp = http_request("GET", endpoint, params=params, headers=self._headers())
            if resp.status_code != 200:
                raise ConnectorError(describe_http_error(resp, self.display_name))
            try:
                data = resp.json()
            except Exception as exc:
                raise ConnectorError(f"返回不是 JSON：{exc}") from exc
            result = data.get("search-results") or {}
            entries = result.get("entry") or []
            if not entries:
                break
            for item in entries:
                if "error" in item:
                    continue
                if item.get("subtype") == "EC":
                    continue
                records.append(make_record(
                    source_id=self.id,
                    source=self.display_name,
                    source_type="api",
                    title=item.get("dc:title"),
                    authors=item.get("dc:creator") or "",
                    year=(item.get("prism:coverDate") or "")[:4],
                    venue=item.get("prism:publicationName") or "",
                    doi=item.get("prism:doi"),
                    url=item.get("prism:url") or "",
                    abstract=item.get("dc:description") or "",
                    keywords=item.get("authkeywords") or "",
                    query_used=q,
                    note=f"EID {item.get('eid', '')}；被引 {item.get('citedby-count', '')}",
                ))
                if len(records) >= max_results:
                    break
            start += self.max_per_page
        return records


class ScienceDirectConnector(_ElsevierBase):
    connector_type = "sciencedirect"
    default_endpoint = "https://api.elsevier.com/content/search/sciencedirect"

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or self.default_endpoint).strip()
        q = self._apply_year(q, year_range)
        records: List[Dict[str, Any]] = []
        start = 0
        while len(records) < max_results and start <= 200:
            params = {
                "query": q,
                "count": min(self.max_per_page, max_results - len(records)),
                "start": start,
            }
            params.update(self.extra_params())
            resp = http_request("GET", endpoint, params=params, headers=self._headers())
            if resp.status_code != 200:
                raise ConnectorError(describe_http_error(resp, self.display_name))
            try:
                data = resp.json()
            except Exception as exc:
                raise ConnectorError(f"返回不是 JSON：{exc}") from exc
            result = data.get("search-results") or {}
            entries = result.get("entry") or []
            if not entries:
                break
            for item in entries:
                if "error" in item:
                    continue
                author_block = item.get("authors") or {}
                names = []
                if isinstance(author_block, dict):
                    for a in author_block.get("author") or []:
                        if isinstance(a, dict):
                            name = a.get("$") or a.get("name") or ""
                            if name:
                                names.append(name)
                records.append(make_record(
                    source_id=self.id,
                    source=self.display_name,
                    source_type="api",
                    title=item.get("dc:title"),
                    authors="; ".join(names[:20]),
                    year=(item.get("prism:coverDate") or "")[:4],
                    venue=item.get("prism:publicationName") or "",
                    doi=item.get("prism:doi"),
                    url=item.get("link") or item.get("prism:url") or "",
                    abstract=item.get("dc:description") or "",
                    keywords=item.get("authkeywords") or "",
                    query_used=q,
                    note=f"PII {item.get('pii', '')}",
                ))
                if len(records) >= max_results:
                    break
            start += self.max_per_page
        return records


class IeeeConnector(BaseConnector):
    connector_type = "ieee"

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or "https://ieeexploreapi.ieee.org/api/v1/search/articles").strip()
        params: Dict[str, Any] = {
            "apikey": self.require("api_key", "API Key"),
            "querytext": q,
            "max_records": min(max(1, max_results), 200),
            "start_record": 1,
            "sort_field": "relevance",
            "sort_order": "desc",
        }
        if year_range and any(year_range):
            start, end = year_range
            if start and end:
                params["ranges"] = f"{start}_{end}_Year"
            elif start:
                params["ranges"] = f"{start}_{datetime.now().year}_Year"
            elif end:
                params["ranges"] = f"1900_{end}_Year"
        params.update(self.extra_params())
        resp = http_request("GET", endpoint, params=params, headers=self.extra_headers())
        if resp.status_code != 200:
            raise ConnectorError(describe_http_error(resp, self.display_name))
        try:
            data = resp.json()
        except Exception as exc:
            raise ConnectorError(f"返回不是 JSON：{exc}") from exc

        records: List[Dict[str, Any]] = []
        for item in (data.get("articles") or [])[:max_results]:
            authors_block = item.get("authors") or {}
            names = []
            if isinstance(authors_block, dict):
                for a in authors_block.get("authors") or []:
                    if isinstance(a, dict):
                        name = a.get("full_name") or a.get("name") or ""
                        if name:
                            names.append(name)
            records.append(make_record(
                source_id=self.id,
                source=self.display_name,
                source_type="api",
                title=item.get("title"),
                authors="; ".join(names[:20]),
                year=item.get("publication_year"),
                venue=item.get("publication_title") or "",
                doi=item.get("doi"),
                url=item.get("html_url") or item.get("pdf_url") or "",
                abstract=item.get("abstract") or "",
                keywords="; ".join([
                    str(v) for v in (item.get("index_terms") or {}).get("author_terms", [])
                ] if isinstance(item.get("index_terms"), dict) else []),
                query_used=q,
                note=f"IEEE 文献号 {item.get('article_number', '')}",
            ))
        return records


class SpringerConnector(BaseConnector):
    connector_type = "springer"

    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        endpoint = str(self.cfg.get("endpoint") or "https://api.springernature.com/meta/v2/json").strip()
        params: Dict[str, Any] = {
            "q": q,
            "api_key": self.require("api_key", "API Key"),
            "p": min(max(1, max_results), 100),
            "s": 1,
        }
        params.update(self.extra_params())
        resp = http_request("GET", endpoint, params=params, headers=self.extra_headers())
        if resp.status_code != 200:
            raise ConnectorError(describe_http_error(resp, self.display_name))
        try:
            data = resp.json()
        except Exception as exc:
            raise ConnectorError(f"返回不是 JSON：{exc}") from exc

        records: List[Dict[str, Any]] = []
        for item in (data.get("records") or [])[:max_results]:
            creators = [
                c.get("creator") or c.get("name") or ""
                for c in (item.get("creators") or [])
                if isinstance(c, dict)
            ]
            urls = item.get("url") or []
            url = ""
            if isinstance(urls, list) and urls:
                first = urls[0]
                url = first.get("value") if isinstance(first, dict) else str(first)
            elif isinstance(urls, str):
                url = urls
            records.append(make_record(
                source_id=self.id,
                source=self.display_name,
                source_type="api",
                title=item.get("title"),
                authors="; ".join([c for c in creators if c][:20]),
                year=(item.get("publicationDate") or "")[:4],
                venue=item.get("publicationName") or "",
                doi=item.get("doi"),
                url=url,
                abstract=item.get("abstract") or "",
                keywords="; ".join(item.get("keyword") or []) if isinstance(item.get("keyword"), list) else "",
                query_used=q,
                note=f"类型 {item.get('genre', '')}",
            ))
        return records


class TemplateConnector(BaseConnector):
    """
    通用模板 connector：既可走「已确认的 API + JSON 字段映射」，
    也可走「Session + URL 模板 + HTML CSS 映射」。

    这是自定义文献库与无公开 API 站点的统一实现，不针对任何具体站点写死逻辑。
    """

    connector_type = "template"

    # -- URL 构造 ---------------------------------------------------------
    def _build_url(self, query: str, page: int = 1,
                   year_range: Optional[Tuple[Optional[int], Optional[int]]] = None) -> str:
        tpl = str(self.cfg.get("url_template") or "").strip()
        endpoint = str(self.cfg.get("endpoint") or "").strip()
        if tpl:
            start, end = (year_range or (None, None))
            return (
                tpl.replace("{query}", quote_plus(query))
                .replace("{query_raw}", query)
                .replace("{page}", str(page))
                .replace("{year_start}", str(start or ""))
                .replace("{year_end}", str(end or ""))
            )
        if endpoint:
            return endpoint
        raise ConnectorError(
            f"{self.display_name} 未配置搜索地址：请在「文献库配置」页填写搜索 URL 模板"
            "（例如 https://example.org/search?q={query}）或 API endpoint。"
        )

    def _auth_mode(self) -> str:
        mode = str(self.cfg.get("auth_mode") or "header").strip().lower()
        return mode if mode in ("header", "bearer", "query", "none") else "header"

    def _headers(self) -> Dict[str, str]:
        headers = {"Accept": "application/json, text/html;q=0.9", "User-Agent": USER_AGENT}
        cookie = str(self.cfg.get("cookie") or "").strip()
        if cookie:
            headers["Cookie"] = cookie
        mode = self._auth_mode()
        api_key = str(self.cfg.get("api_key") or "").strip()
        if api_key:
            if mode == "bearer":
                headers["Authorization"] = f"Bearer {api_key}"
            elif mode == "query":
                pass  # 由 search() 以查询参数形式带上
            elif mode != "none":
                header_name = str(self.cfg.get("key_header") or "X-API-Key").strip() or "X-API-Key"
                headers[header_name] = api_key
        headers.update(self.extra_headers())
        return headers

    # -- 检索 -------------------------------------------------------------
    def search(self, query, max_results=20, year_range=None):
        q = (query or "").strip()
        if not q:
            raise ConnectorError("查询式为空。")
        pages = max(1, min(int(self.cfg.get("pages") or 1), 5))
        method = str(self.cfg.get("method") or "GET").upper()
        response_type = str(self.cfg.get("response_type") or "json").lower()
        query_param = str(self.cfg.get("query_param") or "q").strip() or "q"

        collected: List[Dict[str, Any]] = []
        last_error = ""
        for page in range(1, pages + 1):
            url = self._build_url(q, page, year_range)
            params: Dict[str, Any] = {}
            body = None
            extra = self.extra_params()
            # URL 模板里没有 {query} 占位符时，把查询式放进请求参数/请求体，避免"空检索"
            tpl = str(self.cfg.get("url_template") or "")
            needs_param = ("{query}" not in tpl)
            if method == "GET":
                if needs_param:
                    params[query_param] = q
                if page > 1 and "{page}" not in tpl:
                    params.setdefault("page", page)
                params.update(extra)
                if self._auth_mode() == "query":
                    api_key = str(self.cfg.get("api_key") or "").strip()
                    if api_key:
                        params[str(self.cfg.get("key_param") or "apikey")] = api_key
            else:
                body = json.loads(json.dumps(extra)) if extra else {}
                if needs_param:
                    body.setdefault(query_param, q)
                if page > 1 and "{page}" not in tpl:
                    body.setdefault("page", page)
                if self._auth_mode() == "query":
                    api_key = str(self.cfg.get("api_key") or "").strip()
                    if api_key:
                        body.setdefault(str(self.cfg.get("key_param") or "apikey"), api_key)
            if response_type == "html" and self.cfg.get("use_playwright"):
                html_text = self._fetch_with_playwright(url)
                page_records = self._parse_html_text(html_text, q, url)
            else:
                resp = http_request(
                    method, url, params=params or None, headers=self._headers(),
                    json_body=body if method != "GET" else None,
                )
                if resp.status_code != 200:
                    last_error = describe_http_error(resp, self.display_name)
                    break
                page_records = (
                    self._parse_json(resp, q) if response_type == "json"
                    else self._parse_html_text(resp.text, q, url)
                )
            collected.extend(page_records)
            if len(collected) >= max_results or not page_records:
                break

        if not collected and last_error:
            raise ConnectorError(last_error)
        if not collected and response_type == "html":
            raise ConnectorError(
                f"{self.display_name} 返回的 HTML 中未解析出任何记录。"
                "常见原因：① Cookie 失效被重定向到登录页；② 页面由 JavaScript 渲染，"
                "requests 拿到空壳；③ CSS 选择器与当前页面结构不符。"
                "请在文献库配置页核对选择器，或启用 Playwright 渲染。"
            )
        return collected[:max_results]

    def _parse_json(self, resp: requests.Response, query: str) -> List[Dict[str, Any]]:
        try:
            data = resp.json()
        except Exception as exc:
            raise ConnectorError(f"响应不是合法 JSON：{exc}。若该地址返回的是网页，请把响应类型改为 HTML。") from exc

        fm = self.cfg.get("field_map") or {}
        if not isinstance(fm, dict) or not fm:
            raise ConnectorError(f"{self.display_name} 未配置 JSON 字段映射（field_map）。")

        container_path = fm.get("container") or "$"
        nodes = json_path_get(data, container_path, default=None)
        if nodes is None:
            raise ConnectorError(
                f"JSON 路径 {container_path} 没有取到数据，请检查 field_map.container 是否与返回结构一致。"
            )
        if not isinstance(nodes, list):
            nodes = [nodes]

        records: List[Dict[str, Any]] = []
        for node in nodes:
            if not isinstance(node, (dict, list)):
                continue
            values = {}
            for field_name in ("title", "authors", "year", "venue", "doi", "url", "abstract", "keywords"):
                path = fm.get(field_name)
                if path:
                    values[field_name] = json_path_get(node, path, default="")
            title = as_text(values.get("title"))
            if not title and not values.get("doi") and not values.get("url"):
                continue
            records.append(make_record(
                source_id=self.id,
                source=self.display_name,
                source_type=self.source_type,
                title=title,
                authors=as_text(values.get("authors")),
                year=as_text(values.get("year")),
                venue=as_text(values.get("venue")),
                doi=as_text(values.get("doi")),
                url=as_text(values.get("url")),
                abstract=as_text(values.get("abstract")),
                keywords=as_text(values.get("keywords")),
                query_used=query,
                note="模板 API 解析",
            ))
        return records

    def _fetch_with_playwright(self, url: str) -> str:
        """
        可选的浏览器渲染通道。默认不启用，也不依赖任何云端浏览器服务。
        仅当用户在本机自行安装 playwright 并勾选时才走这里。
        """
        try:
            from playwright.sync_api import sync_playwright  # type: ignore
        except Exception as exc:  # pragma: no cover
            raise ConnectorError(
                "已勾选 Playwright 渲染，但本机没有安装 playwright。"
                "请先执行 pip install playwright 与 playwright install chromium，"
                "或取消勾选改用请求直取。"
            ) from exc
        cookie = str(self.cfg.get("cookie") or "").strip()
        try:
            with sync_playwright() as play:  # pragma: no cover
                browser = play.chromium.launch(headless=True)
                context = browser.new_context(user_agent=USER_AGENT)
                if cookie:
                    header_pairs = {}
                    for part in cookie.split(";"):
                        if "=" in part:
                            key, value = part.split("=", 1)
                            header_pairs[key.strip()] = value.strip()
                    if header_pairs:
                        context.add_cookies([
                            {"name": k, "value": v, "domain": urlparse(url).hostname or "",
                             "path": "/"}
                            for k, v in header_pairs.items()
                        ])
                page = context.new_page()
                page.goto(url, timeout=45000, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)
                html_text = page.content()
                browser.close()
                return html_text
        except Exception as exc:  # pragma: no cover
            raise ConnectorError(f"Playwright 渲染失败：{exc.__class__.__name__}: {exc}") from exc

    def _parse_html_text(self, html_text: str, query: str, base_url: str) -> List[Dict[str, Any]]:
        if not BS4_AVAILABLE:
            raise ConnectorError("未安装 beautifulsoup4，无法解析 HTML。请先 pip install beautifulsoup4 lxml。")
        css = self.cfg.get("css_map") or {}
        if not isinstance(css, dict) or not css.get("container"):
            raise ConnectorError(f"{self.display_name} 未配置 CSS 解析规则（css_map.container）。")

        soup = BeautifulSoup(html_text, "lxml")
        nodes = soup.select(css["container"])
        if not nodes:
            raise ConnectorError(
                f"CSS 选择器 '{css['container']}' 未匹配到任何节点。"
                "可能是页面由 JavaScript 渲染（此时请启用 Playwright），或站点结构与模板不符。"
            )

        def pick(node, selector: str, attr: Optional[str] = None) -> str:
            if not selector:
                return ""
            try:
                el = node.select_one(selector)
            except Exception:
                return ""
            if not el:
                return ""
            if attr:
                return str(el.get(attr) or "")
            return el.get_text(" ", strip=True)

        records: List[Dict[str, Any]] = []
        for node in nodes:
            title = pick(node, css.get("title", ""))
            link = pick(node, css.get("link", "") or css.get("title", ""), attr="href")
            if link:
                link = urljoin(base_url, link)
            if not title and not link:
                continue
            records.append(make_record(
                source_id=self.id,
                source=self.display_name,
                source_type="session",
                title=title,
                authors=pick(node, css.get("authors", "")),
                year=pick(node, css.get("year", "")),
                venue=pick(node, css.get("venue", "")),
                doi=pick(node, css.get("doi", "")),
                url=link,
                abstract=pick(node, css.get("abstract", "")),
                keywords=pick(node, css.get("keywords", "")),
                query_used=query,
                note="Session 页面解析，字段完整度取决于页面结构",
            ))
        return records

    # -- 测试连接 ---------------------------------------------------------
    def test_connection(self) -> Tuple[bool, str]:
        mode = str(self.cfg.get("response_type") or "json").lower()
        if mode == "html" and not str(self.cfg.get("url_template") or "").strip():
            return False, "未填写搜索 URL 模板。"
        if mode == "json" and not (self.cfg.get("url_template") or self.cfg.get("endpoint")):
            return False, (
                "未填写 API endpoint 或搜索 URL 模板。该库没有可直接使用的公开接口，"
                "请向机构索取接入地址后填入。"
            )
        return super().test_connection()


CONNECTOR_REGISTRY: Dict[str, type] = {
    "openalex": OpenAlexConnector,
    "crossref": CrossrefConnector,
    "wos": WosConnector,
    "scopus": ScopusConnector,
    "sciencedirect": ScienceDirectConnector,
    "ieee": IeeeConnector,
    "springer": SpringerConnector,
    "template": TemplateConnector,
}


def build_connector(cfg: Dict[str, Any]) -> Optional[BaseConnector]:
    cls = CONNECTOR_REGISTRY.get(str(cfg.get("connector") or "template"))
    if cls is None:
        return None
    return cls(cfg)


def all_source_configs(cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    items = list(cfg.get("sources", {}).values()) + list(cfg.get("custom_sources", []))
    return items


def source_status(source: Dict[str, Any]) -> Tuple[str, str]:
    """返回 (状态标签, 颜色提示)。"""
    test = source.get("test") or {}
    if not source.get("enabled"):
        return "已禁用", "⚪"
    configured = bool(
        str(source.get("api_key") or "").strip()
        or str(source.get("cookie") or "").strip()
        or str(source.get("url_template") or "").strip()
        or source.get("id") in ("openalex", "crossref")
    )
    if not configured:
        return "未配置", "🟡"
    if test.get("ok") is True:
        return "测试通过", "🟢"
    if test.get("ok") is False:
        return "测试失败", "🔴"
    return "已配置未测试", "🔵"


# ---------------------------------------------------------------------------
# 去重
# ---------------------------------------------------------------------------

def record_identity_keys(rec: Dict[str, Any]) -> List[str]:
    """
    一条记录可以同时拥有多个身份键（DOI / 规范化标题 / URL）。
    去重时命中任意一个即视为同一篇文献——因为同一个库返回 DOI、
    另一个库可能只给标题，只看单一键会漏合并。
    """
    keys: List[str] = []
    doi = normalize_doi(rec.get("doi"))
    if doi:
        keys.append(f"doi:{doi}")
    title = normalize_title(rec.get("title"))
    # 过短的标题（如 Editorial、Preface）使用标题键容易误合并，长度不足时不参与
    if title and len(title) >= 8:
        keys.append(f"title:{title[:120]}")
    url = str(rec.get("url") or "").strip().lower()
    if url:
        keys.append(f"url:{url[:160]}")
    return keys


def dedupe_records(records: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], int]:
    """多键并查去重：DOI / 标题 / URL 任意命中即合并，合并后新键也注册回同一实体。"""
    index: Dict[str, Dict[str, Any]] = {}
    ordered: List[Dict[str, Any]] = []
    duplicates = 0

    for rec in records:
        keys = record_identity_keys(rec)
        target: Optional[Dict[str, Any]] = None
        for key in keys:
            if key in index:
                target = index[key]
                break

        if target is None:
            rec["dedup_key"] = keys[0] if keys else f"anon:{len(ordered)}"
            if not keys:
                keys = [rec["dedup_key"]]
            ordered.append(rec)
            for key in keys:
                index[key] = rec
            continue

        duplicates += 1
        for field in ("abstract", "doi", "url", "venue", "keywords", "authors", "year"):
            if not target.get(field) and rec.get(field):
                target[field] = rec[field]
        # 合并后把这一条的新键也指向同一实体，便于链式合并
        for key in keys:
            index.setdefault(key, target)
        note = str(target.get("note") or "")
        source_name = str(rec.get("source") or "")
        if source_name and source_name not in note:
            target["note"] = f"{note}；重复来源：{source_name}".strip("；")

    return ordered, duplicates


# ---------------------------------------------------------------------------
# 启发式打分（LLM 不可用时的兜底）
# ---------------------------------------------------------------------------

def _tokenize(text: str) -> List[str]:
    text = str(text or "").lower()
    return re.findall(r"[a-z0-9]{2,}|[\u4e00-\u9fff]{2,}", text)


def heuristic_score(rec: Dict[str, Any], plan: QueryPlan) -> float:
    query_terms = set(
        _tokenize(" ".join(plan.keywords_en + plan.keywords_zh + plan.synonyms + [plan.topic]))
    )
    if not query_terms:
        return 5.0
    text = " ".join([str(rec.get("title") or ""), str(rec.get("abstract") or ""),
                     str(rec.get("keywords") or "")])
    doc_terms = set(_tokenize(text))
    if not doc_terms:
        return 3.0
    overlap = len(query_terms & doc_terms) / max(1, len(query_terms))
    score = 2.0 + overlap * 6.0
    year = to_int(rec.get("year"))
    if year:
        current = datetime.now().year
        if year >= current - 3:
            score += 1.5
        elif year >= current - 8:
            score += 0.8
    if str(rec.get("abstract") or "").strip():
        score += 0.3
    return round(min(10.0, score), 2)


# ---------------------------------------------------------------------------
# LLM 重排与摘要
# ---------------------------------------------------------------------------

RERANK_SYSTEM_PROMPT = """你是学术文献筛选专家。给定研究主题和一批候选文献，逐条判断它们与主题的相关程度。

评分标准（0-10 分）：
- 9-10：核心相关，主题直接命中，方法或结论可直接引用
- 6-8：较为相关，主题有实质交集
- 3-5：弱相关，仅共享部分术语或领域
- 0-2：不相关或仅字面匹配

只输出 JSON 数组，不要代码块，不要解释。格式：
[{"i": 1, "score": 8.5, "summary": "一句话中文摘要（不超过 60 字）", "reason": "一句话说明相关性（不超过 40 字）"}]

要求：
- 必须为输入的每一条文献都给出结果，i 与输入编号一一对应。
- summary 只描述论文做了什么，不要评价；reason 只说明为什么相关。
- 不要编造论文中不存在的内容；信息不足时按标题判断，并在 reason 中说明"仅据标题判断"。
"""


def llm_rerank(records: List[Dict[str, Any]], plan: QueryPlan, llm: LLMClient,
               top_n: int = 80, batch_size: int = 8,
               log: Callable[[str], None] = lambda m: None,
               enable_summary: bool = True) -> None:
    """原地修改 records：写入 score / llm_note。失败时保留启发式分数。"""
    for rec in records:
        if not rec.get("score"):
            rec["score"] = heuristic_score(rec, plan)

    if llm is None or not llm.is_configured():
        log("AI API 未配置，跳过 LLM 重排，使用启发式分数。")
        return

    targets = records[: max(1, top_n)]
    total_batches = (len(targets) + batch_size - 1) // batch_size
    log(f"开始 LLM 重排：{len(targets)} 条候选，共 {total_batches} 批。")

    for batch_index in range(total_batches):
        chunk = targets[batch_index * batch_size: (batch_index + 1) * batch_size]
        lines = []
        for idx, rec in enumerate(chunk, 1):
            abstract = str(rec.get("abstract") or "")[:700]
            lines.append(
                f"[{idx}] 标题：{rec.get('title', '')}\n"
                f"    年份：{rec.get('year', '')}；来源：{rec.get('venue', '')}\n"
                f"    摘要：{abstract or '（无摘要）'}"
            )
        user_prompt = (
            f"研究主题：{plan.topic}\n"
            f"核心概念：{', '.join(plan.keywords_en or plan.keywords_zh)}\n\n"
            "候选文献：\n" + "\n".join(lines) + "\n\n请输出 JSON 数组。"
        )
        try:
            raw = llm.chat(
                [
                    {"role": "system", "content": RERANK_SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.1,
                max_tokens=min(2000, 220 * len(chunk) + 200),
            )
            parsed = parse_json_object(raw)
            if not isinstance(parsed, list):
                log(f"第 {batch_index + 1}/{total_batches} 批返回无法解析，保留启发式分数。")
                continue
            for item in parsed:
                if not isinstance(item, dict):
                    continue
                try:
                    idx = int(item.get("i"))
                except Exception:
                    continue
                if not 1 <= idx <= len(chunk):
                    continue
                rec = chunk[idx - 1]
                try:
                    rec["score"] = round(float(item.get("score", rec.get("score", 0))), 2)
                except Exception:
                    pass
                summary = clean_text(item.get("summary"), 200)
                reason = clean_text(item.get("reason"), 120)
                if enable_summary and (summary or reason):
                    rec["llm_note"] = f"{summary}" + (f"（相关性：{reason}）" if reason else "")
                elif reason:
                    rec["llm_note"] = f"（相关性：{reason}）"
            log(f"第 {batch_index + 1}/{total_batches} 批评分完成。")
        except LLMError as exc:
            log(f"第 {batch_index + 1}/{total_batches} 批失败：{exc}（保留启发式分数）")
        except Exception as exc:  # noqa: BLE001
            log(f"第 {batch_index + 1}/{total_batches} 批异常：{exc.__class__.__name__}: {exc}")


# ---------------------------------------------------------------------------
# 主流程编排
# ---------------------------------------------------------------------------

@dataclass
class SourceStat:
    source_id: str
    display_name: str
    ok: bool
    count: int
    message: str
    elapsed: float


def run_pipeline(
    user_query: str,
    cfg: Dict[str, Any],
    *,
    year_start: Optional[int] = None,
    year_end: Optional[int] = None,
    max_results: int = 30,
    enable_rerank: bool = True,
    enable_summary: bool = True,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    """
    完整流水线：意图解析 -> 并发/串行检索 -> 去重 -> LLM 重排 -> 汇总。
    任一步骤失败都被隔离，返回结构里带上每库状态。
    """
    log = log or (lambda m: None)
    progress_cb = progress_cb or (lambda p, t: None)

    llm_cfg = cfg.get("llm") or {}
    llm: Optional[LLMClient] = LLMClient(llm_cfg) if llm_cfg else None
    if llm is not None and not llm.is_configured():
        llm = None

    # --- 1. 意图解析 -----------------------------------------------------
    progress_cb(0.05, "解析检索意图")
    plan = build_query_plan(user_query, llm, log=log)
    if year_start:
        plan.year_start = year_start
    if year_end:
        plan.year_end = year_end
    year_range = (plan.year_start, plan.year_end)
    log(f"主题：{plan.topic or '（未识别）'}")
    log(f"英文关键词：{', '.join(plan.keywords_en) or '无'}")
    log(f"中文关键词：{', '.join(plan.keywords_zh) or '无'}")
    log(f"年份范围：{year_range[0] or '不限'} - {year_range[1] or '不限'}")

    # --- 2. 逐库检索 -----------------------------------------------------
    enabled = [s for s in all_source_configs(cfg) if s.get("enabled")]
    if not enabled:
        raise ConnectorError("没有任何已启用的文献库，请先在「文献库配置」页启用至少一个。")

    stats: List[SourceStat] = []
    all_records: List[Dict[str, Any]] = []
    step = 0.10
    span = 0.55 / max(1, len(enabled))

    for source in enabled:
        step += span
        name = source.get("display_name") or source.get("id")
        progress_cb(min(step, 0.70), f"正在检索 {name}")
        log(f"→ 检索 {name} …")
        connector = build_connector(source)
        started = time.perf_counter()
        if connector is None:
            message = f"未知的 connector 类型：{source.get('connector')}"
            stats.append(SourceStat(source.get("id", ""), name, False, 0, message, 0.0))
            log(f"  ✗ {message}")
            continue
        try:
            query_text = plan.query_for(str(source.get("id")))
            if not connector.supports_boolean:
                # 该库只做词袋匹配：改送纯关键词，避免把 AND/OR 当检索词
                plain = " ".join(plan.keywords_en[:6]) or " ".join(plan.keywords_zh[:6])
                if plain.strip():
                    query_text = plain
            records = connector.search(
                query_text,
                max_results=max_results,
                year_range=year_range,
            )
            elapsed = time.perf_counter() - started
            for rec in records:
                rec["source_id"] = source.get("id")
                rec["source"] = name
            all_records.extend(records)
            stats.append(SourceStat(source.get("id", ""), name, True, len(records), "", elapsed))
            log(f"  ✓ {name} 返回 {len(records)} 条（{elapsed:.1f}s）")
        except ConnectorError as exc:
            elapsed = time.perf_counter() - started
            stats.append(SourceStat(source.get("id", ""), name, False, 0, str(exc), elapsed))
            log(f"  ✗ {name} 失败：{exc}")
        except Exception as exc:  # noqa: BLE001
            elapsed = time.perf_counter() - started
            message = f"{exc.__class__.__name__}: {exc}"
            stats.append(SourceStat(source.get("id", ""), name, False, 0, message, elapsed))
            log(f"  ✗ {name} 异常：{message}")

    # --- 3. 去重 ---------------------------------------------------------
    progress_cb(0.72, "去重中")
    log(f"原始记录 {len(all_records)} 条，开始去重…")
    unique_records, dup_count = dedupe_records(all_records)
    log(f"去重完成：合并 {dup_count} 条重复，保留 {len(unique_records)} 条。")

    for rec in unique_records:
        if not rec.get("score"):
            rec["score"] = heuristic_score(rec, plan)

    # --- 4. 重排 / 摘要 --------------------------------------------------
    if enable_rerank and unique_records:
        progress_cb(0.80, "LLM 重排中")
        llm_rerank(
            unique_records, plan, llm,
            top_n=int((cfg.get("settings") or {}).get("rerank_top_n", 80)),
            log=log, enable_summary=enable_summary,
        )
    else:
        log("已跳过 LLM 重排，仅保留启发式分数。")

    def score_value(item: Dict[str, Any]) -> float:
        try:
            return float(item.get("score") or 0)
        except Exception:
            return 0.0

    unique_records.sort(key=lambda r: (score_value(r), to_int(r.get("year")) or 0), reverse=True)

    progress_cb(1.0, "完成")
    return {
        "plan": plan,
        "records": unique_records,
        "stats": stats,
        "duplicates": dup_count,
        "raw_count": len(all_records),
        "finished_at": now_str(),
    }


# ---------------------------------------------------------------------------
# Excel 导出
# ---------------------------------------------------------------------------

def records_to_dataframe(records: List[Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for rec in records:
        row = {}
        for header, key in GRAPH_COLUMNS:
            value = rec.get(key, "")
            if key == "source_type":
                value = {"api": "API", "session": "Session", "manual": "手工"}.get(str(value), value)
            row[header] = value
        rows.append(row)
    return pd.DataFrame(rows, columns=[h for h, _ in GRAPH_COLUMNS])


def dataframe_to_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="检索结果")
        worksheet = writer.sheets["检索结果"]
        widths = {
            "文献库": 22, "标题": 52, "作者": 32, "年份": 8, "期刊/会议": 28,
            "DOI": 30, "链接": 38, "摘要": 60, "关键词": 26, "相关性分数": 10,
            "LLM 摘要/推荐理由": 46, "检索查询": 40, "数据源类型": 12,
            "抓取时间": 20, "去重标识": 34, "原始错误或备注": 30,
        }
        for idx, header in enumerate(df.columns, start=1):
            letter = get_column_letter(idx)
            worksheet.column_dimensions[letter].width = widths.get(header, 18)
        if OPENPYXL_AVAILABLE:
            head_fill = PatternFill("solid", fgColor="EEF2FF")
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="1F2937")
                cell.fill = head_fill
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Streamlit 界面
# ---------------------------------------------------------------------------

def init_state() -> None:
    if "cfg" not in st.session_state:
        st.session_state["cfg"] = load_config()
    st.session_state.setdefault("show_secrets", False)
    st.session_state.setdefault("llm_test", None)
    st.session_state.setdefault("result", None)
    st.session_state.setdefault("run_log", [])


def json_text(value: Any) -> str:
    try:
        return json.dumps(value, ensure_ascii=False, indent=2)
    except Exception:
        return "{}"


def parse_json_text(text: str, label: str) -> Tuple[Optional[Any], str]:
    text = (text or "").strip()
    if not text:
        return {}, ""
    try:
        return json.loads(text), ""
    except Exception as exc:
        return None, f"{label} 不是合法 JSON：{exc}"


def persist(cfg: Dict[str, Any]) -> None:
    ok, message = save_config(cfg)
    if ok:
        st.success(message)
    else:
        st.error(message)


# ------------------------------ 页签 1 -------------------------------------

def render_ai_tab(cfg: Dict[str, Any]) -> None:
    st.subheader("AI API 配置")
    st.caption(
        "兼容任何 OpenAI 风格的 /chat/completions 接口：DeepSeek、Qwen、Moonshot、"
        "Ollama、vLLM、OpenRouter、各类中转网关。程序只依赖这一个协议。"
    )
    llm = cfg["llm"]
    reveal = st.session_state["show_secrets"]

    with st.form("llm_form"):
        col1, col2 = st.columns([3, 2])
        with col1:
            base_url = st.text_input("Base URL", value=llm.get("base_url", ""),
                                     placeholder="https://api.deepseek.com")
            model = st.text_input("Model", value=llm.get("model", ""),
                                  placeholder="deepseek-chat")
        with col2:
            api_key = st.text_input(
                "API Key", value=llm.get("api_key", ""),
                type="default" if reveal else "password",
                placeholder="sk-...",
            )
            temperature = st.slider("Temperature", 0.0, 1.5,
                                    float(llm.get("temperature", 0.2)), 0.05)
        col3, col4 = st.columns(2)
        with col3:
            max_tokens = st.number_input("Max tokens", 128, 32000,
                                         int(llm.get("max_tokens", 2048)), 128)
        with col4:
            timeout = st.number_input("请求超时（秒）", 10, 600,
                                      int(llm.get("timeout", 120)), 10)

        saved = st.form_submit_button("保存配置", use_container_width=True)
        tested = st.form_submit_button("保存并测试连接", use_container_width=True)

    if saved or tested:
        cfg["llm"] = {
            "base_url": base_url.strip(),
            "api_key": api_key.strip(),
            "model": model.strip(),
            "temperature": float(temperature),
            "max_tokens": int(max_tokens),
            "timeout": int(timeout),
        }
        persist(cfg)
        st.session_state["cfg"] = cfg

    if tested:
        with st.spinner("正在调用模型…"):
            client = LLMClient(cfg["llm"])
            ok, message = client.test_connection()
        st.session_state["llm_test"] = (ok, message)

    result = st.session_state.get("llm_test")
    if result:
        ok, message = result
        if ok:
            st.success("连接成功，模型回复：")
            st.code(message, language="text")
        else:
            st.error(f"连接失败：{message}")

    st.divider()
    st.markdown("**当前配置摘要**")
    st.json({
        "base_url": cfg["llm"].get("base_url", ""),
        "model": cfg["llm"].get("model", ""),
        "api_key": mask_secret(cfg["llm"].get("api_key", "")),
        "temperature": cfg["llm"].get("temperature"),
    })


# ------------------------------ 页签 2 -------------------------------------

def render_source_card(source: Dict[str, Any], cfg: Dict[str, Any], index: int) -> None:
    sid = source.get("id", f"src{index}")
    label, dot = source_status(source)
    title = f"{dot} {source.get('display_name', sid)}　·　{label}"
    if not source.get("preset", True):
        title += "　·　自定义"

    with st.expander(title, expanded=False):
        preset = PRESET_BY_ID.get(sid, {})
        notes = source.get("notes") or preset.get("notes") or ""
        if notes:
            st.caption(notes)
        if not source.get("verified", False):
            st.warning(
                "该库没有可直接使用的公开检索 API。下面预填的地址与解析规则只是可编辑模板，"
                "未经验证；请按你的机构实际入口修改，或只启用你有权限的方式。",
                icon="⚠️",
            )

        connector_type = source.get("connector", "template")
        if connector_type in ("openalex", "crossref"):
            auth_label = "免密钥"
        elif str(source.get("api_key") or "").strip():
            auth_label = "API Key"
        elif str(source.get("cookie") or "").strip():
            auth_label = "Cookie / Session"
        else:
            auth_label = "未配置凭据"
        render_mode = "JSON / API" if str(source.get("response_type") or "json") == "json" else "HTML / 页面解析"
        st.caption(
            f"类型：{source.get('source_type', 'api')}　|　认证方式：{auth_label}　|　"
            f"响应形态：{render_mode}　|　connector：{connector_type}"
        )
        is_api_family = connector_type in ("wos", "scopus", "sciencedirect", "ieee", "springer", "openalex", "crossref")
        is_template = connector_type == "template"

        with st.form(f"form_{sid}_{index}"):
            enabled = st.toggle("启用该文献库", value=bool(source.get("enabled")),
                                key=f"en_{sid}_{index}")
            st.markdown("---")

            if connector_type in ("openalex", "crossref"):
                st.info(f"「{source.get('display_name', sid)}」是开放源，无需 API Key，可留空直接使用。")
                endpoint = st.text_input("Endpoint", value=source.get("endpoint", ""), key=f"ep_{sid}_{index}")
                email = st.text_input("联系邮箱（可选，进入 polite pool）",
                                      value=source.get("inst_token", ""), key=f"mail_{sid}_{index}")
                api_key = ""
                inst_token = email
                headers_text = json_text(source.get("extra_headers") or {})
                params_text = json_text(source.get("extra_params") or {})
                url_template = ""
                method = "GET"
                query_param = "q"
                auth_mode = "header"
                key_header = "X-API-Key"
                key_param = "apikey"
                cookie = ""
                response_type = "json"
                field_map_text = "{}"
                css_map_text = "{}"
                pages = 1
                use_playwright = False
                db = ""

            elif is_api_family:
                endpoint = st.text_input("API Endpoint", value=source.get("endpoint", ""),
                                         key=f"ep_{sid}_{index}")
                c1, c2 = st.columns(2)
                with c1:
                    api_key = st.text_input(
                        "API Key", value=source.get("api_key", ""),
                        type="default" if st.session_state["show_secrets"] else "password",
                        key=f"key_{sid}_{index}",
                    )
                with c2:
                    inst_token = st.text_input(
                        "Institutional Token / Email（可选）",
                        value=source.get("inst_token", ""),
                        type="default" if st.session_state["show_secrets"] else "password",
                        key=f"inst_{sid}_{index}",
                    )
                if connector_type == "wos":
                    db = st.text_input("数据库标识（db 参数）", value=source.get("db", "WOS"),
                                       key=f"db_{sid}_{index}")
                else:
                    db = source.get("db", "")
                headers_text = st.text_area("额外 Headers（JSON）",
                                            value=json_text(source.get("extra_headers") or {}),
                                            height=80, key=f"hdr_{sid}_{index}")
                params_text = st.text_area("额外请求参数（JSON）",
                                           value=json_text(source.get("extra_params") or {}),
                                           height=80, key=f"par_{sid}_{index}")
                url_template = ""
                method = "GET"
                query_param = "q"
                auth_mode = "header"
                key_header = "X-API-Key"
                key_param = "apikey"
                cookie = ""
                response_type = "json"
                field_map_text = "{}"
                css_map_text = "{}"
                pages = 1
                use_playwright = False

            else:  # template
                mode = st.radio(
                    "接入方式",
                    ["Session / 网页模板", "API / JSON 模板"],
                    index=0 if source.get("source_type") == "session" else 1,
                    horizontal=True,
                    key=f"mode_{sid}_{index}",
                )
                response_type = "html" if mode.startswith("Session") else "json"
                url_template = st.text_input(
                    "搜索 URL 模板（可用占位符 {query} {page} {year_start} {year_end}）",
                    value=source.get("url_template", ""), key=f"url_{sid}_{index}",
                )
                endpoint = st.text_input(
                    "API Endpoint（JSON 模板模式下使用；留空则用上面的 URL 模板）",
                    value=source.get("endpoint", ""), key=f"ep_{sid}_{index}",
                )
                c1, c2, c3 = st.columns(3)
                with c1:
                    method = st.selectbox("请求方法", ["GET", "POST"],
                                          index=0 if (source.get("method") or "GET") == "GET" else 1,
                                          key=f"m_{sid}_{index}")
                with c2:
                    pages = st.number_input("抓取页数", 1, 5, int(source.get("pages") or 1),
                                            key=f"pg_{sid}_{index}")
                with c3:
                    query_param = st.text_input(
                        "查询参数名", value=source.get("query_param", "q"),
                        help="URL 模板里没有 {query} 占位符时，查询式会放进这个请求参数（或 POST 请求体的同名字段）。",
                        key=f"qp_{sid}_{index}",
                    )
                c3, c4 = st.columns(2)
                with c3:
                    api_key = st.text_input(
                        "API Key（可选）", value=source.get("api_key", ""),
                        type="default" if st.session_state["show_secrets"] else "password",
                        key=f"key_{sid}_{index}",
                    )
                with c4:
                    auth_mode_label = st.selectbox(
                        "API Key 传递方式",
                        ["请求头（自定义字段名）", "Authorization: Bearer", "查询参数", "不使用"],
                        index={"header": 0, "bearer": 1, "query": 2, "none": 3}.get(
                            str(source.get("auth_mode") or "header"), 0),
                        key=f"am_{sid}_{index}",
                    )
                auth_mode = {"请求头（自定义字段名）": "header", "Authorization: Bearer": "bearer",
                             "查询参数": "query", "不使用": "none"}[auth_mode_label]
                kc3, kc4 = st.columns(2)
                with kc3:
                    key_header = st.text_input(
                        "Key 请求头字段名（header 模式）",
                        value=source.get("key_header", "X-API-Key"), key=f"kh_{sid}_{index}",
                    )
                with kc4:
                    key_param = st.text_input(
                        "Key 查询参数名（query 模式）",
                        value=source.get("key_param", "apikey"), key=f"kp_{sid}_{index}",
                    )
                cookie = st.text_input(
                    "Cookie（Session 模式）", value=source.get("cookie", ""),
                    type="default" if st.session_state["show_secrets"] else "password",
                    help="从浏览器开发者工具 Network 面板复制整段 Cookie。默认掩码显示。",
                    key=f"ck_{sid}_{index}",
                )
                headers_text = st.text_area("额外 Headers（JSON）",
                                            value=json_text(source.get("extra_headers") or {}),
                                            height=80, key=f"hdr_{sid}_{index}")
                params_text = st.text_area("额外请求参数（JSON）",
                                           value=json_text(source.get("extra_params") or {}),
                                           height=80, key=f"par_{sid}_{index}")
                if response_type == "json":
                    field_map_text = st.text_area(
                        "JSON 字段映射（container + 各字段路径）",
                        value=json_text(source.get("field_map") or {
                            "container": "$.results[*]",
                            "title": "title", "authors": "authors[*].name",
                            "year": "year", "venue": "venue", "doi": "doi",
                            "url": "url", "abstract": "abstract",
                        }),
                        height=180, key=f"fm_{sid}_{index}",
                    )
                    css_map_text = "{}"
                else:
                    css_map_text = st.text_area(
                        "CSS 解析规则（container + 各字段选择器）",
                        value=json_text(source.get("css_map") or {
                            "container": "article",
                            "title": "h2 a", "authors": ".authors",
                            "link": "h2 a",
                        }),
                        height=180, key=f"css_{sid}_{index}",
                    )
                    field_map_text = "{}"
                use_playwright = st.checkbox(
                    "用 Playwright 渲染页面（需本机已安装 playwright，默认关闭）",
                    value=bool(source.get("use_playwright")),
                    help="仅用于必须执行 JavaScript 才能出结果的站点。启用后程序会在本地启动无头 Chromium，"
                         "不经过任何云端浏览器。未安装 playwright 时会给出明确报错。",
                    key=f"pw_{sid}_{index}",
                )
                db = ""
                inst_token = ""
            saved = st.form_submit_button("保存该库配置", use_container_width=True)
            tested = st.form_submit_button("保存并测试连接", use_container_width=True)
            if not source.get("preset", True):
                pass  # 删除按钮放在 form 外，避免与 submit 冲突

        if saved or tested:
            headers_obj, err1 = parse_json_text(headers_text, "Headers")
            params_obj, err2 = parse_json_text(params_text, "请求参数")
            fields_obj, err3 = parse_json_text(field_map_text, "字段映射")
            css_obj, err4 = parse_json_text(css_map_text, "CSS 规则")
            errors = [e for e in (err1, err2, err3, err4) if e]
            if errors:
                for err in errors:
                    st.error(err)
            else:
                source.update({
                    "enabled": bool(enabled),
                    "endpoint": endpoint.strip(),
                    "api_key": api_key.strip(),
                    "inst_token": inst_token.strip(),
                    "db": (db or "").strip(),
                    "extra_headers": headers_obj or {},
                    "extra_params": params_obj or {},
                    "url_template": url_template.strip(),
                    "method": method,
                    "query_param": (query_param or "q").strip(),
                    "auth_mode": auth_mode,
                    "key_header": (key_header or "X-API-Key").strip(),
                    "key_param": (key_param or "apikey").strip(),
                    "cookie": cookie.strip(),
                    "response_type": response_type,
                    "field_map": fields_obj or {},
                    "css_map": css_obj or {},
                    "pages": int(pages),
                    "use_playwright": bool(use_playwright),
                    "source_type": "session" if response_type == "html" else "api",
                })
                persist(cfg)
                if tested:
                    connector = build_connector(source)
                    with st.spinner("正在测试连接…"):
                        ok, message = (False, "connector 构建失败") if connector is None else connector.test_connection()
                    source["test"] = {"ok": ok, "msg": message, "ts": now_str()}
                    persist(cfg)
                    if ok:
                        st.success(message)
                    else:
                        st.error(message)
                st.rerun()

        # 测试状态与删除
        test = source.get("test") or {}
        if test.get("ts"):
            if test.get("ok"):
                st.caption(f"✅ 上次测试 {test.get('ts')}：{test.get('msg')}")
            else:
                st.caption(f"❌ 上次测试 {test.get('ts')}：{test.get('msg')}")

        if not source.get("preset", True):
            if st.button("删除该自定义文献库", key=f"del_{sid}_{index}"):
                cfg["custom_sources"] = [s for s in cfg.get("custom_sources", []) if s.get("id") != sid]
                persist(cfg)
                st.rerun()


def render_sources_tab(cfg: Dict[str, Any]) -> None:
    st.subheader("文献库配置")
    st.caption(
        "先启用你有权限的库。API 优先，Session 作为补位。"
        "凭据只写入本机 config.json，界面默认掩码显示。"
    )

    reveal = st.toggle("显示敏感字段明文（仅本地排查用）", value=st.session_state["show_secrets"])
    st.session_state["show_secrets"] = reveal

    enabled_sources = [s for s in all_source_configs(cfg) if s.get("enabled")]
    st.info(f"当前已启用 {len(enabled_sources)} 个文献库：" +
            ("、".join(s.get("display_name", "") for s in enabled_sources) or "无"))

    st.markdown("### 默认文献库")
    default_ids = [p["id"] for p in PRESETS]
    for idx, sid in enumerate(default_ids):
        source = cfg["sources"].get(sid)
        if source:
            render_source_card(source, cfg, idx)

    st.markdown("### 自定义文献库")
    with st.expander("＋ 添加自定义文献库", expanded=False):
        with st.form("add_custom"):
            name = st.text_input("名称", placeholder="例如：某机构内部知识库")
            c1, c2 = st.columns(2)
            with c1:
                kind = st.selectbox("类型", ["API（JSON 响应）", "Session（HTML 响应）"])
            with c2:
                new_pages = st.number_input("抓取页数", 1, 5, 1)
            url_template = st.text_input(
                "搜索 URL 模板",
                placeholder="https://example.org/search?q={query}&page={page}",
            )
            endpoint = st.text_input("API Endpoint（JSON 模式下可与上二选一）", value="")
            c3, c4 = st.columns(2)
            with c3:
                method = st.selectbox("请求方法", ["GET", "POST"])
            with c4:
                api_key = st.text_input("API Key（可选）", type="password")
            c5, c6, c7 = st.columns(3)
            with c5:
                auth_mode_label = st.selectbox(
                    "API Key 传递方式",
                    ["请求头（自定义字段名）", "Authorization: Bearer", "查询参数", "不使用"],
                    index=0,
                )
            with c6:
                key_header = st.text_input("Key 请求头字段名", value="X-API-Key")
            with c7:
                query_param = st.text_input("查询参数名", value="q")
            auth_mode = {"请求头（自定义字段名）": "header", "Authorization: Bearer": "bearer",
                         "查询参数": "query", "不使用": "none"}[auth_mode_label]
            cookie = st.text_input("Cookie（Session 模式）", type="password",
                                   help="从浏览器开发者工具复制整段 Cookie，默认掩码显示。")
            headers_text = st.text_area("额外 Headers（JSON）", value="{}", height=70)
            params_text = st.text_area("额外请求参数（JSON）", value="{}", height=70)
            if kind.startswith("API"):
                mapping_text = st.text_area(
                    "JSON 字段映射",
                    value=json_text({
                        "container": "$.results[*]",
                        "title": "title", "authors": "authors[*].name",
                        "year": "year", "venue": "venue", "doi": "doi",
                        "url": "url", "abstract": "abstract",
                    }),
                    height=160,
                )
                css_text = "{}"
            else:
                css_text = st.text_area(
                    "CSS 解析规则",
                    value=json_text({
                        "container": "article",
                        "title": "h2 a", "authors": ".authors", "link": "h2 a",
                    }),
                    height=160,
                )
                mapping_text = "{}"
            submitted = st.form_submit_button("添加并启用")

        if submitted:
            if not name.strip():
                st.error("请填写名称。")
            elif not (url_template.strip() or endpoint.strip()):
                st.error("请至少填写搜索 URL 模板或 API Endpoint。")
            else:
                mapping_obj, e1 = parse_json_text(mapping_text, "字段映射")
                css_obj, e2 = parse_json_text(css_text, "CSS 规则")
                headers_obj, e3 = parse_json_text(headers_text, "Headers")
                params_obj, e4 = parse_json_text(params_text, "请求参数")
                errs = [e for e in (e1, e2, e3, e4) if e]
                if errs:
                    for err in errs:
                        st.error(err)
                else:
                    safe_id = re.sub(r"[^a-z0-9_]", "_", name.strip().lower())[:24] or f"custom_{int(time.time())}"
                    custom_id = safe_id
                    existing = {s.get("id") for s in all_source_configs(cfg)}
                    suffix = 1
                    while custom_id in existing:
                        suffix += 1
                        custom_id = f"{safe_id}_{suffix}"
                    cfg.setdefault("custom_sources", []).append({
                        "id": custom_id,
                        "display_name": name.strip(),
                        "connector": "template",
                        "source_type": "session" if kind.startswith("Session") else "api",
                        "preset": False,
                        "enabled": True,
                        "endpoint": endpoint.strip(),
                        "api_key": api_key.strip(),
                        "inst_token": "",
                        "db": "",
                        "extra_headers": headers_obj or {},
                        "extra_params": params_obj or {},
                        "url_template": url_template.strip(),
                        "method": method,
                        "query_param": query_param.strip() or "q",
                        "auth_mode": auth_mode,
                        "key_header": key_header.strip() or "X-API-Key",
                        "key_param": "apikey",
                        "cookie": cookie.strip(),
                        "response_type": "html" if kind.startswith("Session") else "json",
                        "field_map": mapping_obj or {},
                        "css_map": css_obj or {},
                        "pages": int(new_pages),
                        "use_playwright": False,
                        "notes": "用户自定义文献库。",
                        "verified": True,
                        "test": {"ok": None, "msg": "", "ts": ""},
                    })
                    persist(cfg)
                    st.rerun()

    customs = cfg.get("custom_sources") or []
    if customs:
        custom_index = 0
        for source in customs:
            render_source_card(source, cfg, 100 + custom_index)
            custom_index += 1
    else:
        st.caption("尚未添加自定义文献库。")

    st.divider()
    st.markdown("### 全局操作")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("清除所有配置（含密钥）", type="secondary", use_container_width=True):
            st.session_state["confirm_clear"] = True
    with c2:
        if st.button("重新载入磁盘上的 config.json", use_container_width=True):
            st.session_state["cfg"] = load_config()
            st.rerun()

    if st.session_state.get("confirm_clear"):
        st.warning("此操作会删除本机 config.json 中的 AI API 与所有文献库凭据，不可撤销。")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("确认清除", type="primary", use_container_width=True):
                try:
                    if CONFIG_PATH.exists():
                        CONFIG_PATH.unlink()
                except Exception as exc:
                    st.error(f"删除失败：{exc}")
                else:
                    st.session_state["cfg"] = default_config()
                    st.session_state["confirm_clear"] = False
                    st.session_state["result"] = None
                    st.success("已清除所有本地配置。")
                    st.rerun()
        with cc2:
            if st.button("取消", use_container_width=True):
                st.session_state["confirm_clear"] = False
                st.rerun()


# ------------------------------ 页签 3 -------------------------------------

def render_search_tab(cfg: Dict[str, Any]) -> None:
    st.subheader("检索")
    enabled = [s for s in all_source_configs(cfg) if s.get("enabled")]
    if not enabled:
        st.warning("还没有启用任何文献库。请先到「文献库配置」页启用至少一个（推荐先启用免密钥的 OpenAlex）。")

    st.markdown("**已启用文献库**")
    st.write("　".join([f"`{s.get('display_name')}`" for s in enabled]) or "无")

    with st.form("search_form"):
        user_query = st.text_area(
            "研究方向 / 研究问题（自然语言即可）",
            height=130,
            placeholder="例如：我想了解大语言模型在管理科学研究中的应用现状，"
                        "特别是文本挖掘与文献计量结合的范式，重点关注 2020 年以后的研究。",
        )
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            year_start = st.number_input("起始年份（0 = 不限）", 0, 2100, 0, 1)
        with c2:
            year_end = st.number_input("结束年份（0 = 不限）", 0, 2100, 0, 1)
        with c3:
            max_results = st.number_input("每个库最多返回", 1, 200,
                                          int((cfg.get("settings") or {}).get("max_results", 30)), 1)
        with c4:
            st.write("")
            st.write("")
            do_rerank = st.checkbox("启用 LLM 重排", value=bool((cfg.get("settings") or {}).get("enable_rerank", True)))
        do_summary = st.checkbox("生成 LLM 摘要与推荐理由",
                                 value=bool((cfg.get("settings") or {}).get("enable_summary", True)))
        submitted = st.form_submit_button("开始检索", type="primary", use_container_width=True)

    if submitted:
        if not user_query.strip():
            st.error("请先输入研究方向。")
        elif not enabled:
            st.error("没有已启用的文献库。")
        else:
            cfg.setdefault("settings", {})["max_results"] = int(max_results)
            cfg["settings"]["enable_rerank"] = bool(do_rerank)
            cfg["settings"]["enable_summary"] = bool(do_summary)
            persist(cfg)

            progress = st.progress(0.0, text="准备中…")
            log_box = st.empty()
            logs: List[str] = []

            def log(message: str) -> None:
                logs.append(f"[{datetime.now():%H:%M:%S}] {message}")
                log_box.code("\n".join(logs[-400:]), language="text")

            def progress_cb(value: float, text: str) -> None:
                progress.progress(min(max(value, 0.0), 1.0), text=text)

            try:
                result = run_pipeline(
                    user_query.strip(),
                    cfg,
                    year_start=int(year_start) or None,
                    year_end=int(year_end) or None,
                    max_results=int(max_results),
                    enable_rerank=bool(do_rerank),
                    enable_summary=bool(do_summary),
                    progress_cb=progress_cb,
                    log=log,
                )
                st.session_state["result"] = result
                st.session_state["run_log"] = logs
                progress.progress(1.0, text="完成")
                st.success(
                    f"检索完成：原始 {result['raw_count']} 条，去重后 {len(result['records'])} 条。"
                )
            except ConnectorError as exc:
                progress.empty()
                st.error(str(exc))
            except Exception as exc:  # noqa: BLE001
                progress.empty()
                st.exception(exc)

    result = st.session_state.get("result")
    if result:
        st.divider()
        st.markdown("**各库状态**")
        stat_rows = [{
            "文献库": s.display_name,
            "状态": "成功" if s.ok else "失败",
            "返回条数": s.count,
            "耗时(秒)": round(s.elapsed, 1),
            "说明": s.message,
        } for s in result["stats"]]
        st.dataframe(pd.DataFrame(stat_rows), use_container_width=True, hide_index=True)

        with st.expander("查看解析出的检索计划", expanded=False):
            plan: QueryPlan = result["plan"]
            st.json({
                "topic": plan.topic,
                "keywords_zh": plan.keywords_zh,
                "keywords_en": plan.keywords_en,
                "synonyms": plan.synonyms,
                "boolean_queries": plan.boolean_queries,
                "year_range": {"start": plan.year_start, "end": plan.year_end},
                "降级模式": plan.degraded,
            })

        st.markdown("**结果预览（前 10 条）**")
        df = records_to_dataframe(result["records"])
        st.dataframe(df.head(10), use_container_width=True, hide_index=True)
        st.caption("完整结果与 Excel 下载请到「结果与导出」页。")


# ------------------------------ 页签 4 -------------------------------------

def render_export_tab(cfg: Dict[str, Any]) -> None:
    st.subheader("结果与导出")
    result = st.session_state.get("result")
    if not result or not result.get("records"):
        st.info("还没有检索结果。请先到「检索」页执行一次检索。")
        return

    df_all = records_to_dataframe(result["records"])
    st.caption(f"共 {len(df_all)} 条记录，检索完成于 {result.get('finished_at', '')}")

    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        sources = sorted({str(v) for v in df_all["文献库"].tolist() if str(v)})
        chosen = st.multiselect("按来源筛选", sources, default=sources)
    with c2:
        years = [y for y in df_all["年份"].tolist() if str(y).isdigit()]
        if years:
            y_min, y_max = min(int(y) for y in years), max(int(y) for y in years)
            if y_min < y_max:
                selected_years = st.slider("按年份筛选", y_min, y_max, (y_min, y_max))
            else:
                selected_years = (y_min, y_max)
                st.caption(f"结果年份均为 {y_min}")
        else:
            selected_years = None
    with c3:
        sort_by = st.selectbox("排序", ["相关性分数（高→低）", "年份（新→旧）", "年份（旧→新）", "文献库"])

    df = df_all.copy()
    if chosen:
        df = df[df["文献库"].isin(chosen)]
    if selected_years:
        numeric_year = pd.to_numeric(df["年份"], errors="coerce")
        df = df[(numeric_year.isna()) | ((numeric_year >= selected_years[0]) & (numeric_year <= selected_years[1]))]

    if sort_by.startswith("相关性"):
        df["_score"] = pd.to_numeric(df["相关性分数"], errors="coerce").fillna(0)
        df = df.sort_values("_score", ascending=False).drop(columns=["_score"])
    elif sort_by.startswith("年份（新"):
        df = df.assign(_y=pd.to_numeric(df["年份"], errors="coerce")).sort_values("_y", ascending=False, na_position="last").drop(columns=["_y"])
    elif sort_by.startswith("年份（旧"):
        df = df.assign(_y=pd.to_numeric(df["年份"], errors="coerce")).sort_values("_y", ascending=True, na_position="last").drop(columns=["_y"])
    else:
        df = df.sort_values("文献库")

    st.dataframe(df, use_container_width=True, hide_index=True, height=520)

    try:
        excel_bytes = dataframe_to_excel(df)
        filename = f"paper_seek_plus_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.download_button(
            "下载 Excel（当前筛选结果）",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption(f"文件包含 {len(df)} 行 × {len(df.columns)} 列，已冻结表头并自动换行。")
    except Exception as exc:  # noqa: BLE001
        st.error(f"生成 Excel 失败：{exc}")


# ------------------------------ 侧边栏 -------------------------------------

def render_sidebar(cfg: Dict[str, Any]) -> None:
    with st.sidebar:
        st.markdown("## PaperSeek+")
        st.caption("自然语言 → 多库文献检索 → LLM 重排 → Excel")

        llm_ok = bool(cfg["llm"].get("base_url") and cfg["llm"].get("model"))
        st.markdown(f"AI API：{'🟢 已配置' if llm_ok else '🟡 未配置'}")
        enabled = [s for s in all_source_configs(cfg) if s.get("enabled")]
        st.markdown(f"已启用文献库：**{len(enabled)}**")
        if enabled:
            for s in enabled:
                label, dot = source_status(s)
                st.caption(f"{dot} {s.get('display_name')} — {label}")

        st.divider()
        st.markdown("**配置存储位置**")
        st.code(str(CONFIG_PATH), language="text")
        st.caption(
            "仅保存在本机。Session/Cookie 属于敏感信息，请只在本人设备使用；"
            "使用完毕后可在「文献库配置」页一键清除。"
        )
        st.divider()
        with st.expander("使用提示", expanded=False):
            st.markdown(
                "1. 先在「AI API 配置」页填 Base URL / API Key / Model 并测试。\n"
                "2. 到「文献库配置」页启用至少一个库；没有密钥时可先用 OpenAlex 验证流程。\n"
                "3. 在「检索」页输入研究方向，观察日志与各库状态。\n"
                "4. 到「结果与导出」页筛选并下载 Excel。"
            )
        st.divider()
        st.caption("本工具仅用于访问你有合法权限的资源，不提供登录绕过、验证码识别或付费墙突破能力。")


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="PaperSeek+ 文献检索工作台",
        page_icon="📚",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_state()
    cfg = st.session_state["cfg"]

    st.title("📚 PaperSeek+ 文献检索工作台")
    st.caption(
        "大模型负责理解问题与生成检索意图，程序负责调用各文献库，"
        "模型再参与重排与摘要。所有密钥只存本机。"
    )

    render_sidebar(cfg)

    tab_ai, tab_src, tab_search, tab_export = st.tabs(
        ["① AI API 配置", "② 文献库配置", "③ 检索", "④ 结果与导出"]
    )
    with tab_ai:
        render_ai_tab(cfg)
    with tab_src:
        render_sources_tab(cfg)
    with tab_search:
        render_search_tab(cfg)
    with tab_export:
        render_export_tab(cfg)


if __name__ == "__main__":
    main()
