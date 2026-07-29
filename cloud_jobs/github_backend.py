from __future__ import annotations

import json
import mimetypes
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping
from urllib.parse import quote

import requests

from .bundle import (
    FAILURE_ASSET_NAME,
    INPUT_ASSET_NAME,
    STATUS_ASSET_NAME,
)

# Accept results created by both worker generations. These are intentionally
# independent of cloud_jobs.bundle so a partially updated repository can still
# start and recover existing jobs.
RESULT_ASSET_NAME = "suggested_reports_final_combined.zip"
LEGACY_RESULT_ASSET_NAME = "suggested_reports_all_batches.zip"

API_ROOT = "https://api.github.com"
API_VERSION = "2026-03-10"
DEFAULT_WORKFLOW = "process-suggested-cloud.yml"
DEFAULT_REF = "main"
JOB_TAG_PREFIX = "mmw-job-"


class GitHubJobError(RuntimeError):
    """Raised when GitHub rejects a cloud-job request."""


@dataclass(frozen=True)
class GitHubCloudConfig:
    repository: str
    token: str
    workflow: str = DEFAULT_WORKFLOW
    ref: str = DEFAULT_REF

    @classmethod
    def from_mapping(cls, values: Mapping[str, Any]) -> "GitHubCloudConfig":
        repository = str(values.get("repository", "")).strip()
        if not repository:
            owner = str(values.get("owner", "")).strip()
            repo = str(values.get("repo", "")).strip()
            if owner and repo:
                repository = f"{owner}/{repo}"
        token = str(values.get("token", "")).strip()
        workflow = str(values.get("workflow", DEFAULT_WORKFLOW)).strip() or DEFAULT_WORKFLOW
        ref = str(values.get("ref", values.get("branch", DEFAULT_REF))).strip() or DEFAULT_REF

        if repository.count("/") != 1 or any(part.strip() == "" for part in repository.split("/")):
            raise GitHubJobError("The full-processing location is not configured correctly.")
        if not token:
            raise GitHubJobError("The full-processing access setting is missing.")
        return cls(repository=repository, token=token, workflow=workflow, ref=ref)


class GitHubReleaseBackend:
    """Manage private, draft-release-backed jobs through GitHub's REST API."""

    def __init__(self, config: GitHubCloudConfig, timeout: tuple[int, int] = (30, 600)) -> None:
        self.config = config
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/vnd.github+json",
                "Authorization": f"Bearer {config.token}",
                "X-GitHub-Api-Version": API_VERSION,
                "User-Agent": "matching-mapper-world-streamlit",
            }
        )

    @property
    def repository_api(self) -> str:
        return f"{API_ROOT}/repos/{self.config.repository}"

    def _request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        timeout = kwargs.pop("timeout", self.timeout)
        response = self.session.request(method, url, timeout=timeout, **kwargs)
        if response.status_code >= 400:
            detail = response.text[:1500]
            try:
                payload = response.json()
                detail = str(payload.get("message") or detail)
                if payload.get("errors"):
                    detail += f" Errors: {payload['errors']}"
            except (ValueError, TypeError):
                pass
            raise GitHubJobError(f"The processing service could not complete the request ({response.status_code}). Please try again.")
        return response

    def check_connection(self, *, require_private: bool = True) -> dict[str, Any]:
        repository = self._request("GET", self.repository_api).json()
        if require_private and not bool(repository.get("private")):
            raise GitHubJobError(
                "The secure processing area must be private before business files can be uploaded."
            )
        return repository

    def create_job_release(self, *, job_tag: str, requested_by: str = "") -> dict[str, Any]:
        requester = str(requested_by or "").strip()
        requester_line = f"\nRequested by: {requester}" if requester else ""
        payload = {
            "tag_name": job_tag,
            "target_commitish": self.config.ref,
            "name": f"Matching Mapper job {job_tag}",
            "body": (
                "Status: UPLOADING\n"
                "Temporary private draft release used for one Matching Mapper background job."
                f"{requester_line}"
            ),
            "draft": True,
            "prerelease": False,
            "make_latest": "false",
        }
        return self._request("POST", f"{self.repository_api}/releases", json=payload).json()

    def get_release(self, release_id: int) -> dict[str, Any]:
        return self._request("GET", f"{self.repository_api}/releases/{int(release_id)}").json()

    def list_assets(self, release_id: int) -> list[dict[str, Any]]:
        response = self._request(
            "GET",
            f"{self.repository_api}/releases/{int(release_id)}/assets",
            params={"per_page": 100},
        )
        return list(response.json())

    def update_release_body(self, release_id: int, body: str) -> dict[str, Any]:
        return self._request(
            "PATCH",
            f"{self.repository_api}/releases/{int(release_id)}",
            json={"body": str(body)},
        ).json()

    def upload_asset(
        self,
        *,
        release: Mapping[str, Any],
        file_path: str | Path,
        asset_name: str,
        content_type: str | None = None,
    ) -> dict[str, Any]:
        path = Path(file_path)
        if not path.is_file():
            raise FileNotFoundError(path)
        endpoint = str(release["upload_url"]).split("{", 1)[0]
        headers = {
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {self.config.token}",
            "X-GitHub-Api-Version": API_VERSION,
            "User-Agent": "matching-mapper-world-streamlit",
            "Content-Type": content_type or mimetypes.guess_type(asset_name)[0] or "application/octet-stream",
            "Content-Length": str(path.stat().st_size),
        }
        with path.open("rb") as handle:
            response = self.session.post(
                endpoint,
                params={"name": asset_name},
                headers=headers,
                data=handle,
                timeout=(30, 1800),
            )
        if response.status_code >= 400:
            detail = response.text[:1500]
            try:
                detail = str(response.json().get("message") or detail)
            except (ValueError, TypeError):
                pass
            raise GitHubJobError(f"A file could not be transferred for processing ({response.status_code}). Please try again.")
        return response.json()

    def dispatch(self, *, release_id: int, job_tag: str) -> None:
        workflow = quote(self.config.workflow, safe="")
        self._request(
            "POST",
            f"{self.repository_api}/actions/workflows/{workflow}/dispatches",
            json={
                "ref": self.config.ref,
                "inputs": {"release_id": str(int(release_id)), "job_tag": str(job_tag)},
            },
        )

    def submit_job(self, *, job_tag: str, bundle_path: str | Path, requested_by: str = "") -> dict[str, Any]:
        release: dict[str, Any] | None = None
        try:
            release = self.create_job_release(job_tag=job_tag, requested_by=requested_by)
            self.upload_asset(
                release=release,
                file_path=bundle_path,
                asset_name=INPUT_ASSET_NAME,
                content_type="application/zip",
            )
            self.update_release_body(
                int(release["id"]),
                "Status: QUEUED\nFiles received. Waiting for processing to start.",
            )
            self.dispatch(release_id=int(release["id"]), job_tag=job_tag)
            return release
        except Exception:
            if release is not None:
                try:
                    self.delete_job(int(release["id"]), tag_name=str(release.get("tag_name", job_tag)))
                except Exception:
                    pass
            raise

    def list_recent_jobs(self, limit: int = 20) -> list[dict[str, Any]]:
        response = self._request(
            "GET",
            f"{self.repository_api}/releases",
            params={"per_page": min(max(int(limit) * 2, 20), 100)},
        )
        jobs = [
            release
            for release in response.json()
            if str(release.get("tag_name", "")).startswith(JOB_TAG_PREFIX)
        ]
        return jobs[: max(1, int(limit))]

    def find_job_by_tag(self, tag_name: str) -> dict[str, Any]:
        target = str(tag_name).strip()
        for release in self.list_recent_jobs(limit=50):
            if str(release.get("tag_name", "")) == target:
                return release
        raise GitHubJobError("The requested Matching Mapper job was not found among recent private jobs.")

    def list_active_jobs(self, limit: int = 30) -> list[dict[str, Any]]:
        """Return draft jobs whose release body still reports an active state."""
        active_markers = (
            "Status: UPLOADING",
            "Status: QUEUED",
            "Status: STARTING",
            "Status: RUNNING",
        )
        return [
            release
            for release in self.list_recent_jobs(limit=limit)
            if any(marker in str(release.get("body", "")) for marker in active_markers)
        ]

    @staticmethod
    def assets_by_name(assets: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
        return {str(asset.get("name")): asset for asset in assets}

    def download_asset_bytes(self, asset: Mapping[str, Any], *, max_bytes: int = 2_000_000) -> bytes:
        headers = dict(self.session.headers)
        headers["Accept"] = "application/octet-stream"
        with self.session.get(
            str(asset["url"]),
            headers=headers,
            stream=True,
            timeout=(30, 300),
        ) as response:
            if response.status_code >= 400:
                raise GitHubJobError(f"Unable to download the requested file ({response.status_code}).")
            chunks: list[bytes] = []
            total = 0
            for chunk in response.iter_content(chunk_size=64 * 1024):
                if not chunk:
                    continue
                total += len(chunk)
                if total > max_bytes:
                    raise GitHubJobError("The status asset is unexpectedly large.")
                chunks.append(chunk)
            return b"".join(chunks)

    def read_json_asset(self, asset: Mapping[str, Any]) -> dict[str, Any]:
        try:
            return json.loads(self.download_asset_bytes(asset).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise GitHubJobError(f"The request status could not be read.") from exc

    def job_snapshot(self, release_id: int) -> dict[str, Any]:
        release = self.get_release(release_id)
        assets = self.list_assets(release_id)
        by_name = self.assets_by_name(assets)
        status: dict[str, Any]
        if STATUS_ASSET_NAME in by_name:
            status = self.read_json_asset(by_name[STATUS_ASSET_NAME])
        else:
            body = str(release.get("body", ""))
            if "Status: RUNNING" in body:
                state = "running"
                message = "Your request is being processed."
            elif "Status: FAILED" in body:
                state = "failed"
                message = "Processing could not be completed."
            else:
                state = "queued"
                message = "Waiting for processing to start."
            status = {"state": state, "message": message}

        result_present = RESULT_ASSET_NAME in by_name or LEGACY_RESULT_ASSET_NAME in by_name
        if result_present and status.get("state") in {"queued", "starting", "running"}:
            status = {**status, "state": "completed", "message": "The combined result is ready."}
        if FAILURE_ASSET_NAME in by_name and not result_present:
            status = {**status, "state": "failed"}

        return {
            "release": release,
            "assets": assets,
            "assets_by_name": by_name,
            "status": status,
        }

    def temporary_asset_download_url(self, asset: Mapping[str, Any]) -> str | None:
        """Return GitHub's short-lived signed redirect URL when the API supplies one."""
        headers = dict(self.session.headers)
        headers["Accept"] = "application/octet-stream"
        with self.session.get(
            str(asset["url"]),
            headers=headers,
            allow_redirects=False,
            stream=True,
            timeout=(30, 300),
        ) as response:
            if response.status_code in {301, 302, 303, 307, 308}:
                return response.headers.get("Location")
            if response.status_code >= 400:
                raise GitHubJobError(f"Unable to create a download link ({response.status_code}).")
            # Some GitHub/API configurations stream a 200 response directly.
            # Do not buffer it here; the caller will use the disk-streaming fallback.
            return None

    def download_asset_to_path(self, asset: Mapping[str, Any], destination: str | Path) -> Path:
        target = Path(destination)
        target.parent.mkdir(parents=True, exist_ok=True)
        headers = dict(self.session.headers)
        headers["Accept"] = "application/octet-stream"
        with self.session.get(
            str(asset["url"]),
            headers=headers,
            stream=True,
            timeout=(30, 1800),
        ) as response:
            if response.status_code >= 400:
                raise GitHubJobError(f"Unable to download the requested file ({response.status_code}).")
            with target.open("wb") as handle:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        handle.write(chunk)
        return target

    def delete_job(self, release_id: int, *, tag_name: str = "") -> None:
        release = self.get_release(release_id)
        tag = str(tag_name or release.get("tag_name", "")).strip()
        self._request("DELETE", f"{self.repository_api}/releases/{int(release_id)}")
        if tag:
            encoded = quote(tag, safe="")
            response = self.session.delete(
                f"{self.repository_api}/git/refs/tags/{encoded}",
                timeout=self.timeout,
            )
            if response.status_code not in {204, 404, 422}:
                raise GitHubJobError(
                    f"The temporary request files were deleted, but final cleanup could not be completed ({response.status_code})."
                )
