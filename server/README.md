# Spell STT proxy

Spell's cloud speech mode sends audio to this proxy; the proxy holds the
Google credentials, so no Google key ships with Spell. Weak PCs use cloud
mode automatically (first-start CPU check in `src/cpu_check.rs`).

## Run

```bash
python3 -m venv venv && venv/bin/pip install -r requirements.txt
export SPELL_APP_KEY='<long random secret>'          # required
export SPELL_GOOGLE_API_KEY='<google api key>'        # or a service account:
# export GOOGLE_APPLICATION_CREDENTIALS=/etc/spell/sa.json
# export SPELL_GCP_PROJECT=<project id>              # (EU, chirp_2 model)
venv/bin/uvicorn stt_proxy:app --host 127.0.0.1 --port 8790
```

Put it behind the existing HTTPS reverse proxy (nginx/Caddy), e.g.
`https://<server>/spell-stt/` → `http://127.0.0.1:8790/`. Audio must only
travel over HTTPS.

Example systemd unit (`/etc/systemd/system/spell-stt-proxy.service`):

```ini
[Unit]
Description=Spell STT proxy
After=network.target

[Service]
WorkingDirectory=/opt/spell-stt-proxy
EnvironmentFile=/etc/spell/stt-proxy.env
ExecStart=/opt/spell-stt-proxy/venv/bin/uvicorn stt_proxy:app --host 127.0.0.1 --port 8790
Restart=always
User=spell

[Install]
WantedBy=multi-user.target
```

`/etc/spell/stt-proxy.env` holds `SPELL_APP_KEY=...` and the Google
credential, readable only by the service user.

## Point Spell at it

Update the `stt_cloud.json` that Spell downloads from S3:

```json
{"stt_cloud_url": "https://<server>/spell-stt", "stt_cloud_key": "<SPELL_APP_KEY>", "stt_google_api_key": ""}
```

Empty `stt_google_api_key` is important: then no Google key is on pupil
machines any more. Afterwards, rotate the Google API key that was shipped
before (it has been on every client).

## API

- `POST /transcribe`, body 16 kHz mono 16-bit WAV, headers `X-Spell-Key`,
  `X-Spell-Lang` (`no|nb|nn|en`) → `{"text": "..."}`.
  Errors: 403 bad key, 400 bad audio/language, 413 > ~70 s, 429 rate limit,
  502 provider error (details never echoed).
- `GET /health` → `{"ok": true, "backend": "google-v1|google-v2"}`.

## Limits

`SPELL_RATE_PER_MIN` (default 300) per client IP — a school often shares one
public IP. Behind a reverse proxy all requests come from 127.0.0.1, so set
the limit at the reverse proxy instead, or raise this one.

## Tests

```bash
venv/bin/pip install pytest && venv/bin/python -m pytest -q test_stt_proxy.py
```
