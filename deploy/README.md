# Deployment — Artifact Generator

Productionized POC → systemd service. Multi-format pptx/docx/xlsx, binds
`172.17.0.1:9191` (Dify-reachable via the docker0 bridge gateway; no longer
`0.0.0.0`).

## Standard reference
- KIWIAGENCY-PLATFORM-OPS-STANDARD-V1 §5 (process management: systemd,
  `Restart=on-failure`, `RestartSec≥5`, enable on boot), §6 (deploy hygiene).
- KIWIAGENCY-DIFY-FIRST-KPRX-THIN-PROXY-STANDARD-V1 (generator independent of KPRX).

## Files
- `app/server.py` — service (ThreadingHTTPServer, configurable bind/output/templates).
- `templates/synthetic_template.{pptx,docx,xlsx}` — default templates for `default` mode.
- `deploy/artifact-generator.service` — systemd unit (User=kiwi, Restart=on-failure, RestartSec=5).

## Configuration (env, all optional)
| Var | Default | Purpose |
|---|---|---|
| `AG_BIND_HOST` | `172.17.0.1` | bind host (NOT 0.0.0.0) |
| `AG_OUTPUT_DIR` | `<repo>/output` | where generated files are stored |
| `AG_TEMPLATES_DIR` | `<repo>/templates` | default-template root |
| `AG_PUBLIC_BASE` | `http://172.17.0.1:9191` | base used in contract `download_url` |

Port defaults to **9191** (matches the POC; unchanged external contract).

## Install (Phase D — run only with card authorization)
```bash
# 1. deploy code
sudo mkdir -p /opt/kiwiai/apps/artifact_generator
sudo cp -a . /opt/kiwiai/apps/artifact_generator/
# 2. venv
cd /opt/kiwiai/apps/artifact_generator
sudo -u kiwi python3 -m venv venv
sudo -u kiwi venv/bin/pip install -r requirements.txt
# 3. unit
sudo cp deploy/artifact-generator.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now artifact-generator   # the single authorized start
# 4. verify
curl -s http://172.17.0.1:9191/health
```

## Contract (unchanged from POC — KPRX/Dify consume it unchanged)
`POST /generate/{fmt}/metadata` →
```json
{ "artifact": {
    "provider": "external_tool",
    "provider_file_id": "<sha[:12]>",
    "download_url": "http://172.17.0.1:9191/download/<fmt>/<fname>",
    "filename": "<title>.<fmt>",
    "mime_type": "...",
    "artifact_type": "<fmt>",
    "size": <bytes>,
    "checksum": "sha256:<full>"
}}
```
`GET /download/{fmt}/{fname}` serves the stored file. Unknown format → HTTP 400.
All three formats are OOXML (ZIP) containers — magic bytes `PK\x03\x04`.
