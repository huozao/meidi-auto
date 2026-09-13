"""每日邮件产物归档：支持本地目录和坚果云 WebDAV。"""

from __future__ import annotations

import base64
import os
import re
import shutil
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


TZ_SHANGHAI = ZoneInfo("Asia/Shanghai")
DATE_RE = re.compile(r"(?:总库存|物料情况)[_-]?(?P<date>\d{8})[_-]?(?P<time>\d{6})")


def _archive_stamp(source: Path) -> datetime:
    match = DATE_RE.search(source.stem)
    if match:
        try:
            return datetime.strptime(match.group("date") + match.group("time"), "%Y%m%d%H%M%S").replace(tzinfo=TZ_SHANGHAI)
        except ValueError:
            pass
    return datetime.fromtimestamp(source.stat().st_mtime, TZ_SHANGHAI)


def _webdav_url(base: str, *parts: str) -> str:
    base = base.rstrip("/") + "/"
    encoded = "/".join(urllib.parse.quote(part.strip("/"), safe="") for part in parts if part.strip("/"))
    return base + encoded


class WebDavStore:
    def __init__(self, base_url: str, user: str, password: str) -> None:
        self.base_url = base_url.rstrip("/") + "/"
        token = base64.b64encode(f"{user}:{password}".encode()).decode()
        self.auth_header = f"Basic {token}"

    def _request(self, method: str, url: str, data: bytes | None = None) -> None:
        request = urllib.request.Request(url, data=data, method=method, headers={"Authorization": self.auth_header})
        try:
            with urllib.request.urlopen(request, timeout=60):
                return
        except urllib.error.HTTPError as exc:
            # Nutstore 对已存在目录返回 405；父目录已存在时 MKCOL 可能返回 409。
            if method == "MKCOL" and exc.code in {405, 409}:
                return
            raise RuntimeError(f"WebDAV {method} {url} 失败: HTTP {exc.code}") from exc

    def list_files(self, remote_dir: str) -> list[str]:
        url = _webdav_url(self.base_url, *remote_dir.strip("/").split("/"))
        request = urllib.request.Request(url, method="PROPFIND", headers={"Authorization": self.auth_header, "Depth": "1"})
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                root = ET.fromstring(response.read())
        except Exception as exc:
            raise RuntimeError(f"WebDAV PROPFIND {remote_dir} 失败: {exc}") from exc
        names: list[str] = []
        for href in root.findall(".//{DAV:}href"):
            value = urllib.parse.unquote(href.text or "").rstrip("/")
            name = value.rsplit("/", 1)[-1]
            if name and name != remote_dir.rstrip("/").rsplit("/", 1)[-1]:
                names.append(name)
        return sorted(set(names))

    def get(self, remote_path: str) -> bytes:
        url = _webdav_url(self.base_url, *remote_path.strip("/").split("/"))
        request = urllib.request.Request(url, method="GET", headers={"Authorization": self.auth_header})
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                return response.read()
        except Exception as exc:
            raise RuntimeError(f"WebDAV GET {remote_path} 失败: {exc}") from exc

    def ensure_dir(self, remote_dir: str) -> None:
        parts = [part for part in remote_dir.strip("/").split("/") if part]
        current: list[str] = []
        for part in parts:
            current.append(part)
            self._request("MKCOL", _webdav_url(self.base_url, *current))

    def put(self, remote_path: str, payload: bytes) -> None:
        parent = "/".join(remote_path.strip("/").split("/")[:-1])
        if parent:
            self.ensure_dir(parent)
        self._request("PUT", _webdav_url(self.base_url, *remote_path.strip("/").split("/")), payload)


def archive_daily_outputs(data_dir: Path) -> list[Path]:
    """把当日最终 Excel、HTML 和图片复制到 YYYYMM 目录。

    未配置归档后端时返回空列表并跳过，保证旧环境的每日邮件链路不受影响。
    """
    candidates = []
    for pattern in ("*总库存*.xlsx", "output.html", "*美的*.png"):
        files = sorted(data_dir.glob(pattern), key=lambda path: path.stat().st_mtime)
        if files:
            candidates.append(files[-1])
    if not candidates:
        raise FileNotFoundError(f"数据目录没有可归档的邮件产物: {data_dir}")

    stamp = _archive_stamp(next(path for path in candidates if path.suffix.lower() == ".xlsx"))
    month_dir = stamp.strftime("%Y%m")
    local_root = os.getenv("DAILY_ATTACHMENT_ARCHIVE_DIR", "").strip()
    dav_base = os.getenv("NUTSTORE_WEBDAV_URL", "").strip()
    dav_user = os.getenv("NUTSTORE_WEBDAV_USER", "").strip()
    dav_password = os.getenv("NUTSTORE_WEBDAV_APP_PASSWORD", "").strip()
    dav_root = os.getenv("NUTSTORE_REMOTE_DAILY_ARCHIVE_DIR", "").strip()
    if not local_root and not (dav_base and dav_user and dav_password and dav_root):
        print("ℹ️ 未配置 DAILY_ATTACHMENT_ARCHIVE_DIR 或坚果云 WebDAV，跳过每日附件归档。")
        return []

    archived: list[Path] = []
    if local_root:
        local_dir = Path(local_root).expanduser() / month_dir
        local_dir.mkdir(parents=True, exist_ok=True)
        for source in candidates:
            target = local_dir / source.name
            shutil.copy2(source, target)
            archived.append(target)
        print(f"✅ 本地附件归档: {local_dir} ({len(candidates)} 个文件)")

    if dav_base and dav_user and dav_password and dav_root:
        store = WebDavStore(dav_base, dav_user, dav_password)
        remote_dir = f"{dav_root.rstrip('/')}/{month_dir}"
        for source in candidates:
            store.put(f"{remote_dir}/{source.name}", source.read_bytes())
        print(f"✅ 坚果云附件归档: {remote_dir} ({len(candidates)} 个文件)")

    return archived
