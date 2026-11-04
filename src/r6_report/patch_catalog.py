"""Parse Huiji Wiki patch index and structured operator-change templates."""

import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from html.parser import HTMLParser
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
from urllib.parse import urljoin

from .sources import PatchChange, PatchRecord


HUIJI_BASE_URL = "https://r6s.huijiwiki.com"
PATCH_INDEX_URL = HUIJI_BASE_URL + "/wiki/更新补丁总表"
_PATCH_PATTERN = re.compile(r"^(Y\d+S\d+(?:\.\d+)*)", re.IGNORECASE)
_SEASON_PATTERN = re.compile(r"^(Y\d+S\d+)", re.IGNORECASE)
_ISO_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_OFFICIAL_URL_PATTERN = re.compile(
    r"\|来源\s*=\s*\[(https://[^\s\]]+)",
    re.IGNORECASE,
)

CURRENT_DIRECTION_OVERRIDES = {
    ("Y11S2.1", "Dokkaebi"): "削弱",
    ("Y11S2.1", "Thorn"): "削弱",
    ("Y11S2.1", "Solis"): "削弱",
    ("Y11S2.1", "Solid Snake"): "混合",
    ("Y11S2.2", "Dokkaebi"): "混合",
    ("Y11S2.2", "Jäger"): "混合",
}


class PatchCatalogError(ValueError):
    """Raised when Huiji patch metadata cannot be parsed safely."""


@dataclass(frozen=True)
class PatchIndexEntry:
    patch: str
    season: str
    season_name: str
    released: date
    wiki_url: str
    wiki_title: str


@dataclass
class _Cell:
    text: str
    href: Optional[str]


class _PatchTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: List[List[_Cell]] = []
        self._row: Optional[List[_Cell]] = None
        self._cell_text: Optional[List[str]] = None
        self._cell_href: Optional[str] = None

    def handle_starttag(
        self, tag: str, attrs: Sequence[Tuple[str, Optional[str]]]
    ) -> None:
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell_text = []
            self._cell_href = None
        elif tag == "a" and self._cell_text is not None:
            attributes = dict(attrs)
            self._cell_href = attributes.get("href")

    def handle_data(self, data: str) -> None:
        if self._cell_text is not None:
            self._cell_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in ("td", "th") and self._row is not None and self._cell_text is not None:
            text = " ".join("".join(self._cell_text).split())
            self._row.append(_Cell(text=text, href=self._cell_href))
            self._cell_text = None
            self._cell_href = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def parse_patch_index_html(html: str) -> Tuple[PatchIndexEntry, ...]:
    if not isinstance(html, str) or not html.strip():
        raise PatchCatalogError("patch index HTML is empty")
    parser = _PatchTableParser()
    try:
        parser.feed(html)
        parser.close()
    except Exception as error:
        raise PatchCatalogError("invalid patch index HTML: %s" % error) from error

    entries = []
    seen = set()
    for row in parser.rows:
        if len(row) < 3 or row[0].text == "所属赛季":
            continue
        patch_match = _PATCH_PATTERN.match(row[1].text)
        if patch_match is None:
            continue
        patch = patch_match.group(1).upper()
        if patch in seen:
            raise PatchCatalogError("duplicate patch index row: %s" % patch)
        if not _ISO_DATE_PATTERN.fullmatch(row[2].text):
            raise PatchCatalogError(
                "patch index date must use YYYY-MM-DD: %s" % row[2].text
            )
        try:
            released = date.fromisoformat(row[2].text)
        except ValueError as error:
            raise PatchCatalogError("invalid patch index date: %s" % row[2].text) from error
        season_match = _SEASON_PATTERN.match(patch)
        if season_match is None:
            raise PatchCatalogError("cannot derive season from patch: %s" % patch)
        href = row[1].href
        if not isinstance(href, str) or not href:
            raise PatchCatalogError("patch index row has no link: %s" % patch)
        wiki_url = urljoin(HUIJI_BASE_URL, href)
        if not wiki_url.startswith(HUIJI_BASE_URL + "/"):
            raise PatchCatalogError("patch index link is not a Huiji URL: %s" % wiki_url)
        title = row[1].text
        entries.append(
            PatchIndexEntry(
                patch=patch,
                season=season_match.group(1).upper(),
                season_name=row[0].text,
                released=released,
                wiki_url=wiki_url,
                wiki_title=title,
            )
        )
        seen.add(patch)
    if not entries:
        raise PatchCatalogError("patch index contains no patch rows")
    entries.sort(key=lambda entry: (entry.released, entry.patch))
    return tuple(entries)


def select_patch_interval(
    entries: Iterable[PatchIndexEntry],
    lower: date,
    upper: date,
) -> Tuple[PatchIndexEntry, ...]:
    if lower > upper:
        raise PatchCatalogError("patch interval lower date is after upper date")
    selected = [
        entry
        for entry in entries
        if lower < entry.released <= upper
    ]
    selected.sort(key=lambda entry: (entry.released, entry.patch))
    return tuple(selected)


def parse_patch_wikitext(
    entry: PatchIndexEntry,
    wikitext: str,
    operator_names: Iterable[str],
) -> PatchRecord:
    if not isinstance(wikitext, str) or not wikitext.strip():
        raise PatchCatalogError("patch wikitext is empty: %s" % entry.patch)
    official_match = _OFFICIAL_URL_PATTERN.search(wikitext)
    if official_match is None:
        raise PatchCatalogError("patch has no official HTTPS source: %s" % entry.patch)
    official_url = official_match.group(1)

    names_by_token: Dict[str, str] = {}
    for name in operator_names:
        if not isinstance(name, str) or not name.strip():
            raise PatchCatalogError("operator name list contains an empty value")
        token = _operator_token(name)
        if token in names_by_token and names_by_token[token] != name:
            raise PatchCatalogError("operator name tokens collide: %s" % name)
        names_by_token[token] = name

    template = _extract_balanced_template(wikitext, "干员改动")
    changes: List[PatchChange] = []
    if template is not None:
        parts = _split_top_level_template(template)
        if not parts or parts[0].strip() != "干员改动":
            raise PatchCatalogError("invalid operator-change template: %s" % entry.patch)
        arguments = [part.strip() for part in parts[1:]]
        if len(arguments) % 2:
            raise PatchCatalogError(
                "operator-change template has unpaired arguments: %s" % entry.patch
            )
        for index in range(0, len(arguments), 2):
            raw_name = arguments[index]
            detail = _clean_wikitext(arguments[index + 1])
            names = _resolve_operator_names(raw_name, names_by_token)
            if not names:
                raise PatchCatalogError(
                    "unknown operator in %s: %s" % (entry.patch, raw_name)
                )
            if not detail:
                raise PatchCatalogError(
                    "operator change has no detail: %s / %s"
                    % (entry.patch, raw_name)
                )
            for name in names:
                changes.append(
                    PatchChange(
                        direction=classify_direction(
                            entry.patch, name, detail
                        ),
                        subject=name,
                        detail=detail,
                    )
                )

    return PatchRecord(
        patch=entry.patch,
        season=entry.season,
        season_name=entry.season_name,
        released=entry.released,
        wiki_url=entry.wiki_url,
        official_url=official_url,
        changes=tuple(changes),
    )


def classify_direction(patch: str, subject: str, detail: str) -> str:
    override = CURRENT_DIRECTION_OVERRIDES.get((patch, subject))
    if override is not None:
        return override
    if "速度" in detail and ("生命值" in detail or "甲" in detail):
        return "混合"

    positive_patterns = (
        r"(?:冷却|充能|补充|间隔|启动)[^。；\n]{0,18}缩短",
        r"提高",
        r"提升",
        r"增加",
        r"扩大",
        r"新增",
        r"更平顺",
        r"(?:冷却时间|充能时间|补充时间)[^。；\n]{0,12}降低",
        r"降低[^。；\n]{0,12}后坐力",
    )
    negative_patterns = (
        r"冷却[^。；\n]{0,8}延长",
        r"持续时间[^。；\n]{0,12}(?:降低|缩短|减少)",
        r"数量[^。；\n]{0,12}(?:降低|减少)",
        r"作用范围[^。；\n]{0,12}(?:降低|缩小|减少)",
        r"提高[^。；\n]{0,8}后坐力",
        r"移除",
    )
    positive = any(re.search(pattern, detail) for pattern in positive_patterns)
    negative = any(re.search(pattern, detail) for pattern in negative_patterns)
    if positive and negative:
        return "混合"
    if negative:
        return "削弱"
    if positive:
        return "增强"
    return "混合"


def _extract_balanced_template(text: str, template_name: str) -> Optional[str]:
    start = text.find("{{" + template_name)
    if start < 0:
        return None
    depth = 0
    index = start
    while index < len(text) - 1:
        pair = text[index:index + 2]
        if pair == "{{":
            depth += 1
            index += 2
            continue
        if pair == "}}":
            depth -= 1
            index += 2
            if depth == 0:
                return text[start:index]
            continue
        index += 1
    raise PatchCatalogError("unterminated template: %s" % template_name)


def _split_top_level_template(template: str) -> List[str]:
    if not template.startswith("{{") or not template.endswith("}}"):
        raise PatchCatalogError("template is not wrapped in balanced braces")
    content = template[2:-2]
    parts = []
    buffer = []
    template_depth = 0
    link_depth = 0
    index = 0
    while index < len(content):
        pair = content[index:index + 2]
        if pair == "{{":
            template_depth += 1
            buffer.append(pair)
            index += 2
            continue
        if pair == "}}" and template_depth:
            template_depth -= 1
            buffer.append(pair)
            index += 2
            continue
        if pair == "[[":
            link_depth += 1
            buffer.append(pair)
            index += 2
            continue
        if pair == "]]" and link_depth:
            link_depth -= 1
            buffer.append(pair)
            index += 2
            continue
        character = content[index]
        if character == "|" and template_depth == 0 and link_depth == 0:
            parts.append("".join(buffer))
            buffer = []
        else:
            buffer.append(character)
        index += 1
    parts.append("".join(buffer))
    return parts


def _clean_wikitext(text: str) -> str:
    value = re.sub(
        r"\[\[([^\]|]+)\|([^\]]+)\]\]",
        lambda match: match.group(2),
        text,
    )
    value = re.sub(r"\[\[([^\]]+)\]\]", lambda match: match.group(1), value)
    value = re.sub(
        r"\{\{wi\|([^{}|]+)(?:\|[^{}]*)?\}\}",
        lambda match: match.group(1),
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"<[^>]+>", "", value)
    value = value.replace("'''", "").replace("''", "")
    lines = [" ".join(line.split()) for line in value.splitlines()]
    return "；".join(line for line in lines if line)


def _operator_token(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_like = "".join(
        character
        for character in normalized
        if not unicodedata.combining(character)
    )
    return "".join(character for character in ascii_like.upper() if character.isalnum())


def _resolve_operator_names(
    raw_value: str,
    names_by_token: Mapping[str, str],
) -> Tuple[str, ...]:
    """Resolve one operator or an unseparated group of known operators."""
    token = _operator_token(_clean_wikitext(raw_value))
    exact = names_by_token.get(token)
    if exact is not None:
        return (exact,)

    candidates = tuple(
        sorted(
            names_by_token.items(),
            key=lambda item: (-len(item[0]), item[0]),
        )
    )
    memo: Dict[int, Tuple[Tuple[str, ...], ...]] = {}

    def segment(position: int) -> Tuple[Tuple[str, ...], ...]:
        if position == len(token):
            return ((),)
        if position in memo:
            return memo[position]
        solutions = []
        for operator_token, name in candidates:
            if not token.startswith(operator_token, position):
                continue
            for suffix in segment(position + len(operator_token)):
                solutions.append((name,) + suffix)
                if len(solutions) > 1:
                    memo[position] = tuple(solutions)
                    return memo[position]
        memo[position] = tuple(solutions)
        return memo[position]

    solutions = segment(0)
    if len(solutions) != 1 or len(solutions[0]) < 2:
        return ()
    return solutions[0]
