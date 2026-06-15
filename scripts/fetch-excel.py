#!/usr/bin/env python3
"""
Fetch IT asset and spend Excel files from SharePoint via Microsoft Graph API.

Required environment variables:
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, SHAREPOINT_SITE_ID

Optional environment variables:
  SHAREPOINT_DRIVE_ID       – specific drive; auto-detected if omitted
  ASSET_FILE_PATH           – path to asset inventory Excel in SharePoint
  SPEND_FILE_PATH           – path to procurement/app spend Excel in SharePoint
"""

import base64
import os
import sys
import pathlib
import urllib.parse

import msal
import requests

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
SITE_ID = os.environ["SHAREPOINT_SITE_ID"]
DRIVE_ID = os.environ.get("SHAREPOINT_DRIVE_ID", "")

# SharePoint paths – override via env vars to match your folder layout
ASSET_FILE_PATH = os.environ.get(
    "ASSET_FILE_PATH", "Assets Date.xlsx"
)
SPEND_FILE_PATH = os.environ.get(
    "SPEND_FILE_PATH", "Saras Apps & Subscriptions Purchase from Jan 26 .xlsx"
)
PROCUREMENT_FILE_PATH = os.environ.get(
    "PROCUREMENT_FILE_PATH", "Anudeep Excel sheets/IT Budget 2026.xlsx"
)
JOINERS_FILE_PATH = os.environ.get(
    "JOINERS_FILE_PATH", "Anudeep Excel sheets/New Joineings and checklist.xlsx"
)
# Vendor payments workbook — a SharePoint/OneDrive sharing URL works here.
# Best-effort: a failure to download does not fail the run (the report then
# shows "no vendor payments sheet connected"). Use `or` so a blank secret
# falls back to the default share URL.
VENDOR_FILE_PATH = os.environ.get("VENDOR_FILE_PATH") or (
    "https://sarasanalytics0-my.sharepoint.com/:x:/g/personal/"
    "anudeep_kolla_sarasanalytics_com/"
    "IQDQYmpYThX_TqjxmCahWDEWAVPwYQoCmoYrI93oeGnzsLQ?e=NuL9Qc"
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DATA_DIR = pathlib.Path(__file__).resolve().parent.parent / "data"

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_access_token() -> str:
    """Acquire an app-only token via MSAL client credentials."""
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        print(f"Authentication failed: {result.get('error_description', result)}", file=sys.stderr)
        sys.exit(1)
    return result["access_token"]

# ---------------------------------------------------------------------------
# Drive helpers
# ---------------------------------------------------------------------------

def resolve_drive_id(token: str) -> str:
    """Return the default document library drive ID for the site."""
    if DRIVE_ID:
        return DRIVE_ID
    url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    return resp.json()["id"]


def _encode_share_url(url: str) -> str:
    """Convert a SharePoint/OneDrive sharing URL to a Graph 'shares' token."""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + b64


def download_via_share_url(token: str, share_url: str, dest: pathlib.Path) -> None:
    """Download a file given a SharePoint/OneDrive sharing URL."""
    share_token = _encode_share_url(share_url)
    # Resolve to driveItem
    meta_url = f"{GRAPH_BASE}/shares/{share_token}/driveItem"
    resp = requests.get(meta_url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    item = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    # Download content
    content_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    resp = requests.get(
        content_url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=120,
        stream=True,
    )
    resp.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
    size_kb = dest.stat().st_size / 1024
    print(f"  ✓ Downloaded (via share URL) → {dest.name} ({size_kb:.0f} KB)")


def download_file(token: str, drive_id: str, file_path: str, dest: pathlib.Path) -> None:
    """Download a file from SharePoint. Supports both folder paths and sharing URLs."""
    # If the path is a sharing URL, resolve via /shares endpoint
    if file_path.startswith("http://") or file_path.startswith("https://"):
        download_via_share_url(token, file_path, dest)
        return

    # Otherwise treat as a path inside the site drive
    parts = file_path.split("/")
    encoded_parts = [urllib.parse.quote(p) for p in parts]
    encoded = "/".join(encoded_parts)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}:/content"
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=120,
        stream=True,
    )
    resp.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
    size_kb = dest.stat().st_size / 1024
    print(f"  ✓ Downloaded {file_path} → {dest.name} ({size_kb:.0f} KB)")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    print("Authenticating with Microsoft Graph …")
    token = get_access_token()
    drive_id = resolve_drive_id(token)

    files_to_fetch = [
        (ASSET_FILE_PATH, DATA_DIR / "asset_inventory.xlsx"),
        (SPEND_FILE_PATH, DATA_DIR / "spend_tracker.xlsx"),
        (PROCUREMENT_FILE_PATH, DATA_DIR / "procurement_plan.xlsx"),
        (JOINERS_FILE_PATH, DATA_DIR / "joiners_info.xlsx"),
    ]

    print(f"Downloading {len(files_to_fetch)} file(s) from SharePoint …")
    failures = []
    for sp_path, local_path in files_to_fetch:
        try:
            download_file(token, drive_id, sp_path, local_path)
        except requests.HTTPError as exc:
            # Redact tokens but show path (last segment only) so we know which file failed
            filename = sp_path.rsplit("/", 1)[-1]
            print(f"  ✗ Failed to download '{filename}' (path: {sp_path}): {exc}", file=sys.stderr)
            failures.append(sp_path)

    if failures:
        print(f"\n{len(failures)} file(s) failed to download:", file=sys.stderr)
        for f in failures:
            print(f"  - {f}", file=sys.stderr)
        sys.exit(1)

    print("All required files downloaded successfully.")

    # Best-effort: vendor payments workbook (optional — never fail the run).
    if VENDOR_FILE_PATH:
        try:
            download_file(token, drive_id, VENDOR_FILE_PATH, DATA_DIR / "vendor_payments.xlsx")
        except requests.HTTPError as exc:
            print(f"  ✗ Vendor payments file not downloaded ({exc}); report will show "
                  f"'no vendor payments sheet connected'.", file=sys.stderr)


if __name__ == "__main__":
    main()
