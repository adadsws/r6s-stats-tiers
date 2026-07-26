"""Huiji MediaWiki client used by the snapshot collector."""

import json
import shutil
import subprocess
import time
from pathlib import Path
from typing import Callable, Dict, List, Mapping, Optional, Sequence

from . import operator_stats, tier_chart


API_URL = "https://r6s.huijiwiki.com/api.php"


class WikiClientError(RuntimeError):
    """Raised when a MediaWiki API request cannot be completed safely."""


class HuijiClient:
    def __init__(
        self,
        *,
        run_command: Callable[..., object] = subprocess.run,
        which: Callable[[str], Optional[str]] = shutil.which,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self._run_command = run_command
        self._which = which
        self._sleep = sleep

    def fetch_tabx(
        self, title: str, fields: Sequence[str]
    ) -> List[Dict[str, object]]:
        return operator_stats.fetch_tabx_page(title, fields)

    def fetch_parsed_html(self, title: str) -> str:
        result = self._request_json(
            {
                "action": "parse",
                "page": title,
                "prop": "text",
                "format": "json",
                "formatversion": "2",
            }
        )
        try:
            text = result["parse"]["text"]
        except (KeyError, TypeError):
            raise WikiClientError("parsed HTML is missing for %s" % title)
        if not isinstance(text, str) or not text.strip():
            raise WikiClientError("parsed HTML is empty for %s" % title)
        return text

    def fetch_wikitext(self, title: str) -> str:
        result = self._request_json(
            {
                "action": "query",
                "prop": "revisions",
                "rvprop": "content",
                "rvslots": "main",
                "format": "json",
                "formatversion": "2",
                "titles": title,
            }
        )
        try:
            page = result["query"]["pages"][0]
            if page.get("missing") is True:
                raise WikiClientError("Wiki page does not exist: %s" % title)
            content = page["revisions"][0]["slots"]["main"]["content"]
        except WikiClientError:
            raise
        except (KeyError, IndexError, TypeError):
            raise WikiClientError("wikitext is missing for %s" % title)
        if not isinstance(content, str) or not content.strip():
            raise WikiClientError("wikitext is empty for %s" % title)
        return content

    def prepare_operator_icons(self, rows, directory: Path) -> Path:
        return operator_stats.prepare_operator_icons(rows, directory)

    def prepare_gadget_icons(self, items, directory: Path):
        return tier_chart.prepare_gadget_icons(items, directory)

    def _request_json(self, parameters: Mapping[str, str]) -> Mapping[str, object]:
        curl_path = self._which("curl.exe") or self._which("curl")
        if not curl_path:
            raise WikiClientError("curl.exe or curl was not found on PATH")
        command = [
            curl_path,
            "--location",
            "--silent",
            "--show-error",
            "--fail",
            "--max-time",
            "30",
            "--user-agent",
            "r6-report-collector/1.0",
            "--get",
            API_URL,
        ]
        for name, value in parameters.items():
            command.extend(["--data-urlencode", "%s=%s" % (name, value)])

        last_error = "unknown request failure"
        for attempt in range(1, 5):
            try:
                result = self._run_command(
                    command,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=True,
                )
                if not result.stdout.strip():
                    raise ValueError("curl returned empty output")
                document = json.loads(result.stdout)
                if not isinstance(document, dict):
                    raise ValueError("MediaWiki response must be an object")
                if "error" in document:
                    raise WikiClientError(
                        "MediaWiki API error: %s" % document["error"]
                    )
                return document
            except WikiClientError:
                raise
            except subprocess.CalledProcessError as error:
                last_error = "curl failed: %s" % (error.stderr or error)
            except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
                last_error = str(error)
            if attempt < 4:
                self._sleep(float(attempt))
        raise WikiClientError(
            "MediaWiki request failed after 4 attempts: %s" % last_error
        )
