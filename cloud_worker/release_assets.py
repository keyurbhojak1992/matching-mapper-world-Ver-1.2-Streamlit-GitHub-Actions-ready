from __future__ import annotations

import argparse
import mimetypes
import os
import sys
from pathlib import Path
from typing import Any

import requests

API_ROOT = "https://api.github.com"
API_VERSION = "2026-03-10"


class ReleaseAssetError(RuntimeError):
    """Raised when a GitHub release-asset operation fails."""


class ReleaseAssets:
    def __init__(self, repository: str, token: str) -> None:
        if repository.count("/") != 1:
            raise ReleaseAssetError("GITHUB_REPOSITORY must use OWNER/REPOSITORY format.")
        if not token:
            raise ReleaseAssetError("GH_TOKEN is required.")
        self.repository = repository
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/vnd.github+json",
                "Authorization": f"Bearer {token}",
                "X-GitHub-Api-Version": API_VERSION,
                "User-Agent": "matching-mapper-world-worker",
            }
        )

    @property
    def repository_api(self) -> str:
        return f"{API_ROOT}/repos/{self.repository}"

    def _request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        timeout = kwargs.pop("timeout", (30, 600))
        response = self.session.request(method, url, timeout=timeout, **kwargs)
        if response.status_code >= 400:
            detail = response.text[:2000]
            try:
                payload = response.json()
                detail = str(payload.get("message") or detail)
                if payload.get("errors"):
                    detail += f" Errors: {payload['errors']}"
            except (ValueError, TypeError):
                pass
            raise ReleaseAssetError(f"GitHub API request failed ({response.status_code}): {detail}")
        return response

    def get_release(self, release_id: int) -> dict[str, Any]:
        return self._request("GET", f"{self.repository_api}/releases/{int(release_id)}").json()

    def list_assets(self, release_id: int) -> list[dict[str, Any]]:
        return list(
            self._request(
                "GET",
                f"{self.repository_api}/releases/{int(release_id)}/assets",
                params={"per_page": 100},
            ).json()
        )

    def find_asset(self, release_id: int, name: str) -> dict[str, Any] | None:
        target = str(name)
        return next((asset for asset in self.list_assets(release_id) if asset.get("name") == target), None)

    def delete_asset(self, release_id: int, name: str, *, missing_ok: bool = True) -> None:
        asset = self.find_asset(release_id, name)
        if asset is None:
            if missing_ok:
                return
            raise ReleaseAssetError(f"Release asset not found: {name}")
        self._request("DELETE", f"{self.repository_api}/releases/assets/{int(asset['id'])}")

    def download_asset(self, release_id: int, name: str, output: Path) -> Path:
        asset = self.find_asset(release_id, name)
        if asset is None:
            raise ReleaseAssetError(f"Release asset not found: {name}")
        output.parent.mkdir(parents=True, exist_ok=True)
        headers = dict(self.session.headers)
        headers["Accept"] = "application/octet-stream"
        with self.session.get(
            str(asset["url"]),
            headers=headers,
            stream=True,
            timeout=(30, 1800),
        ) as response:
            if response.status_code >= 400:
                raise ReleaseAssetError(f"Unable to download {name} ({response.status_code}).")
            temporary = output.with_suffix(output.suffix + ".part")
            with temporary.open("wb") as handle:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        handle.write(chunk)
            temporary.replace(output)
        return output

    def upload_asset(
        self,
        release_id: int,
        file_path: Path,
        *,
        name: str | None = None,
        replace: bool = True,
    ) -> dict[str, Any]:
        source = Path(file_path)
        if not source.is_file():
            raise FileNotFoundError(source)
        asset_name = str(name or source.name)
        if replace:
            self.delete_asset(release_id, asset_name, missing_ok=True)

        release = self.get_release(release_id)
        endpoint = str(release["upload_url"]).split("{", 1)[0]
        content_type = mimetypes.guess_type(asset_name)[0] or "application/octet-stream"
        headers = {
            "Accept": "application/vnd.github+json",
            "Authorization": self.session.headers["Authorization"],
            "X-GitHub-Api-Version": API_VERSION,
            "User-Agent": "matching-mapper-world-worker",
            "Content-Type": content_type,
            "Content-Length": str(source.stat().st_size),
        }
        with source.open("rb") as handle:
            response = self.session.post(
                endpoint,
                params={"name": asset_name},
                headers=headers,
                data=handle,
                timeout=(30, 1800),
            )
        if response.status_code >= 400:
            detail = response.text[:2000]
            try:
                detail = str(response.json().get("message") or detail)
            except (ValueError, TypeError):
                pass
            raise ReleaseAssetError(f"Unable to upload {asset_name} ({response.status_code}): {detail}")
        return response.json()

    def update_body(self, release_id: int, body: str) -> dict[str, Any]:
        return self._request(
            "PATCH",
            f"{self.repository_api}/releases/{int(release_id)}",
            json={"body": str(body)},
        ).json()


def _client_from_environment() -> ReleaseAssets:
    repository = os.environ.get("GITHUB_REPOSITORY", "").strip()
    token = os.environ.get("GH_TOKEN", os.environ.get("GITHUB_TOKEN", "")).strip()
    return ReleaseAssets(repository, token)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage assets on a private Matching Mapper draft release")
    subparsers = parser.add_subparsers(dest="command", required=True)

    download = subparsers.add_parser("download")
    download.add_argument("--release-id", type=int, required=True)
    download.add_argument("--name", required=True)
    download.add_argument("--output", type=Path, required=True)

    upload = subparsers.add_parser("upload")
    upload.add_argument("--release-id", type=int, required=True)
    upload.add_argument("--file", type=Path, required=True)
    upload.add_argument("--name")
    upload.add_argument("--no-replace", action="store_true")

    delete = subparsers.add_parser("delete")
    delete.add_argument("--release-id", type=int, required=True)
    delete.add_argument("--name", required=True)
    delete.add_argument("--required", action="store_true")

    update = subparsers.add_parser("update-body")
    update.add_argument("--release-id", type=int, required=True)
    body_group = update.add_mutually_exclusive_group(required=True)
    body_group.add_argument("--body")
    body_group.add_argument("--body-file", type=Path)
    return parser


def main() -> int:
    arguments = _build_parser().parse_args()
    client = _client_from_environment()

    if arguments.command == "download":
        client.download_asset(arguments.release_id, arguments.name, arguments.output)
    elif arguments.command == "upload":
        client.upload_asset(
            arguments.release_id,
            arguments.file,
            name=arguments.name,
            replace=not arguments.no_replace,
        )
    elif arguments.command == "delete":
        client.delete_asset(arguments.release_id, arguments.name, missing_ok=not arguments.required)
    elif arguments.command == "update-body":
        body = arguments.body
        if arguments.body_file is not None:
            body = arguments.body_file.read_text(encoding="utf-8")
        client.update_body(arguments.release_id, body or "")
    else:  # pragma: no cover - argparse prevents this
        raise AssertionError(arguments.command)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"release-assets error: {exc}", file=sys.stderr)
        raise
