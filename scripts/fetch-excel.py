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

import os
import sys
import pathlib

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
    "ASSET_FILE_PATH", "IT/Asset Inventory/IT_Asset_Inventory.xlsx"
)
SPEND_FILE_PATH = os.environ.get(
    "SPEND_FILE_PATH", "IT/Procurement/IT_Spend_Tracker.xlsx"
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


def download_file(token: str, drive_id: str, file_path: str, dest: pathlib.Path) -> None:
    """Download a single file from SharePoint to a local path."""
    encoded = file_path.replace("/", ":/") if "/" in file_path else file_path
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
    ]

    print(f"Downloading {len(files_to_fetch)} file(s) from SharePoint …")
    for sp_path, local_path in files_to_fetch:
        try:
            download_file(token, drive_id, sp_path, local_path)
        except requests.HTTPError as exc:
            print(f"  ✗ Failed to download {sp_path}: {exc}", file=sys.stderr)
            sys.exit(1)

    print("All files downloaded successfully.")


if __name__ == "__main__":
    main()
