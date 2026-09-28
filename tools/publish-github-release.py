#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MasterEdit GitHub Release 发布脚本

约定（见 AGENTS.md）：
- 版本号从 src-tauri/tauri.conf.json 读取，不在脚本里硬编码
- 发布说明从 tools/notes-<版本>.txt 读取（与软件中心用同一份文案）
- 资产：安装包（bundle/nsis/MasterEdit_<版本>_x64-setup.exe）与便携版（release/MasterEdit_<版本>_x64.exe）
- 必须先 commit → push main → 打 tag → push tag，再运行本脚本（否则 Release 会把标签指向默认分支 HEAD）
"""

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

REPO = "bigmasterpang/MasterEdit"


def log(msg: str) -> None:
    print(f"[*] {msg}", flush=True)


def get_token() -> str:
    candidate_paths = [
        Path(r"C:\opencode\github_tokens"),
        Path.home() / ".github_token",
    ]
    for cp in candidate_paths:
        if cp.exists():
            try:
                val = cp.read_text(encoding="utf-8").strip()
                if val:
                    return val
            except Exception:
                pass
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if token:
        return token
    raise ValueError("未找到 GitHub Token")


def api_request(url: str, token: str, method: str = "GET", data=None, content_type: str = "application/json"):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "MasterEdit-Release-Script",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if content_type:
        headers["Content-Type"] = content_type

    req = urllib.request.Request(url, headers=headers, method=method)
    if data is not None:
        if isinstance(data, dict):
            req.data = json.dumps(data).encode("utf-8")
        elif isinstance(data, (bytes, bytearray)):
            req.data = data
        else:
            req.data = str(data).encode("utf-8")

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            resp_body = resp.read()
            if not resp_body:
                return None
            return json.loads(resp_body.decode("utf-8"))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} for {method} {url}: {error_body}") from e


def read_version(root_dir: Path) -> str:
    conf_path = root_dir / "src-tauri" / "tauri.conf.json"
    conf = json.loads(conf_path.read_text(encoding="utf-8"))
    version = str(conf.get("version", "")).strip()
    if not version:
        raise ValueError(f"无法从 {conf_path} 读取版本号")
    return version


def read_notes(root_dir: Path, version: str) -> str:
    notes_path = root_dir / "tools" / f"notes-{version}.txt"
    if notes_path.exists():
        text = notes_path.read_text(encoding="utf-8").strip()
        if text:
            return text
    log(f"未找到发布说明 {notes_path}，回退为默认文案")
    return f"### MasterEdit v{version} 更新日志\n\n详见仓库提交记录。"


def main() -> int:
    root_dir = Path(__file__).resolve().parent.parent
    token = get_token()
    version = read_version(root_dir)
    tag_name = f"v{version}"
    release_name = f"MasterEdit v{version}"
    notes = read_notes(root_dir, version)

    release_dir = root_dir / "src-tauri" / "target" / "release"
    installer_path = release_dir / "bundle" / "nsis" / f"MasterEdit_{version}_x64-setup.exe"
    portable_path = release_dir / f"MasterEdit_{version}_x64.exe"

    # 便携版：由 release/MasterEdit.exe 复制成带版本号的名字
    if not portable_path.exists():
        raw_exe = release_dir / "MasterEdit.exe"
        if not raw_exe.exists():
            raise FileNotFoundError(f"未找到便携版可执行文件: {raw_exe}")
        portable_path.write_bytes(raw_exe.read_bytes())

    if not installer_path.exists():
        raise FileNotFoundError(f"未找到安装包: {installer_path}")

    log(f"版本: {version} / 标签: {tag_name}")
    log(f"已就绪资产: {installer_path.name} ({installer_path.stat().st_size / 1024 / 1024:.2f} MB)")
    log(f"已就绪资产: {portable_path.name} ({portable_path.stat().st_size / 1024 / 1024:.2f} MB)")

    # 检查 Release 是否已存在
    release_tag_url = f"https://api.github.com/repos/{REPO}/releases/tags/{tag_name}"
    existing_release = None
    try:
        existing_release = api_request(release_tag_url, token, method="GET")
    except RuntimeError as e:
        if "HTTP 404" not in str(e):
            raise

    if existing_release and "id" in existing_release:
        release_id = existing_release["id"]
        log(f"Release {tag_name} 已存在 (ID: {release_id})，更新标题与发布说明...")
        update_url = f"https://api.github.com/repos/{REPO}/releases/{release_id}"
        release = api_request(update_url, token, method="PATCH", data={
            "name": release_name,
            "body": notes,
        })
    else:
        log(f"正在创建新的 GitHub Release: {tag_name}...")
        create_url = f"https://api.github.com/repos/{REPO}/releases"
        release = api_request(create_url, token, method="POST", data={
            "tag_name": tag_name,
            "name": release_name,
            "body": notes,
            "draft": False,
            "prerelease": False,
        })

    release_id = release["id"]
    upload_url_template = release["upload_url"]
    base_upload_url = upload_url_template.split("{")[0]
    log(f"Release ID: {release_id}")

    # 上传资产（同名资产先删后传）
    current_assets = {a["name"]: a["id"] for a in release.get("assets", [])}
    for asset_path in (installer_path, portable_path):
        filename = asset_path.name
        file_bytes = asset_path.read_bytes()
        if filename in current_assets:
            old_asset_id = current_assets[filename]
            log(f"删除同名旧资产: {filename} (ID: {old_asset_id})...")
            api_request(f"https://api.github.com/repos/{REPO}/releases/assets/{old_asset_id}", token, method="DELETE")

        log(f"上传资产: {filename} ({len(file_bytes) / 1024 / 1024:.2f} MB)...")
        upload_url = f"{base_upload_url}?name={urllib.parse.quote(filename)}"
        uploaded = api_request(upload_url, token, method="POST", data=file_bytes, content_type="application/octet-stream")
        log(f"上传成功: {filename} (Asset ID: {uploaded.get('id')})")

    html_url = release.get("html_url") or f"https://github.com/{REPO}/releases/tag/{tag_name}"
    print("\n==================================================")
    print(f"GITHUB_RELEASE_OK: {html_url}")
    print("==================================================")
    return 0


if __name__ == "__main__":
    sys.exit(main())
