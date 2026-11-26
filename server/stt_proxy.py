#!/usr/bin/env python3
"""Spell STT proxy — server-to-server bridge to Google Speech-to-Text.

Pupil machines POST audio here; THIS server holds the Google credentials.
Clients only know the proxy URL and a rotatable app key, so no Google
credential ever ships in Spell (see src/stt/cloud.rs for the client side).

Credentials (server only, pick one):
    GOOGLE_APPLICATION_CREDENTIALS=/path/sa.json   Speech v2, chirp_2 model,
        SPELL_GCP_PROJECT=<project id>             processed in the EU
                                                   (SPELL_GCP_LOCATION,
                                                   default europe-west4).
    SPELL_GOOGLE_API_KEY=<api key>                 Speech v1 REST
                                                   (latest_long model).
    A service account wins when both are set.

Required:
    SPELL_APP_KEY=<secret>   clients must send it as X-Spell-Key. The proxy
                             refuses to start without it.
Optional:
    SPELL_MAX_CONCURRENT=200 Google calls in flight per worker; beyond that
                             requests get 503 at once instead of queueing.
    SPELL_RATE_PER_MIN=300   max requests per client IP per minute (a whole
                             school often shares one public IP).

Run:
    pip install -r requirements.txt
    uvicorn stt_proxy:app --host 0.0.0.0 --port 8790

Fully async: while one request waits on Google (2-5 s), the worker keeps
serving others, so a single worker handles hundreds of parallel dictations.
The work per request is tiny (validate WAV, forward, return JSON).

Wire format (matches CloudSttEngine):
    POST /transcribe
      Content-Type: audio/wav   (16 kHz mono 16-bit PCM WAV)
      X-Spell-Lang: no|nb|nn|en or nb-NO|nn-NO|en-US (default nb-NO)
      X-Spell-Key:  <app key>
    -> 200 {"text": "..."}
    GET /health -> {"ok": true, "backend": "..."}
"""

import asyncio
import base64
import hmac
import os
import struct
import threading
import time
from collections import defaultdict, deque

import httpx
from fastapi import FastAPI, Header, HTTPException, Request

LANG_MAP = {"no": "nb-NO", "nb": "nb-NO", "nn": "nn-NO", "en": "en-US"}
ALLOWED_LANGS = {"nb-NO", "nn-NO", "en-US"}
# Google sync recognize handles ~60 s; 16 kHz mono s16 = 32 KB/s.
MAX_WAV_BYTES = 70 * 32000 + 44
GOOGLE_V1_URL = "https://speech.googleapis.com/v1/speech:recognize"


class Config:
    def __init__(self, env=os.environ):
        self.app_key = env.get("SPELL_APP_KEY", "")
        self.api_key = env.get("SPELL_GOOGLE_API_KEY", "")
        self.sa_path = env.get("GOOGLE_APPLICATION_CREDENTIALS", "")
        self.project = env.get("SPELL_GCP_PROJECT", "")
        self.location = env.get("SPELL_GCP_LOCATION", "europe-west4")
        self.rate_per_min = int(env.get("SPELL_RATE_PER_MIN", "300"))
        self.max_concurrent = int(env.get("SPELL_MAX_CONCURRENT", "200"))

    @property
    def backend(self) -> str:
        if self.sa_path:
            return "google-v2"
        if self.api_key:
            return "google-v1"
        return "none"

    def check(self):
        if not self.app_key:
            raise RuntimeError("SPELL_APP_KEY must be set")
        if self.backend == "none":
            raise RuntimeError("set GOOGLE_APPLICATION_CREDENTIALS or SPELL_GOOGLE_API_KEY")
        if self.backend == "google-v2" and not self.project:
            raise RuntimeError("SPELL_GCP_PROJECT must be set with a service account")


def parse_wav(wav: bytes) -> bytes:
    """Return the PCM payload of a 16 kHz mono 16-bit PCM WAV, or raise."""
    if len(wav) < 44 or wav[0:4] != b"RIFF" or wav[8:12] != b"WAVE":
        raise ValueError("not a WAV file")
    pos = 12
    fmt_ok = False
    while pos + 8 <= len(wav):
        cid = wav[pos:pos + 4]
        size = struct.unpack("<I", wav[pos + 4:pos + 8])[0]
        body = wav[pos + 8:pos + 8 + size]
        if cid == b"fmt ":
            if len(body) < 16:
                raise ValueError("bad fmt chunk")
            audio_fmt, channels, rate, _, _, bits = struct.unpack("<HHIIHH", body[:16])
            if (audio_fmt, channels, rate, bits) != (1, 1, 16000, 16):
                raise ValueError("expected 16 kHz mono 16-bit PCM")
            fmt_ok = True
        elif cid == b"data":
            if not fmt_ok:
                raise ValueError("data before fmt")
            return body
        pos += 8 + size + (size & 1)
    raise ValueError("no data chunk")


class RateLimiter:
    """Sliding one-minute window per client IP."""

    def __init__(self, per_min: int):
        self.per_min = per_min
        self.hits = defaultdict(deque)
        self.lock = threading.Lock()

    def allow(self, key: str, now: float = None) -> bool:
        now = time.monotonic() if now is None else now
        with self.lock:
            q = self.hits[key]
            while q and now - q[0] > 60:
                q.popleft()
            if len(q) >= self.per_min:
                return False
            q.append(now)
            return True


_http: httpx.AsyncClient = None


def http() -> httpx.AsyncClient:
    """One pooled async client for all Google v1 calls (reuses TLS)."""
    global _http
    if _http is None:
        _http = httpx.AsyncClient(
            timeout=30,
            limits=httpx.Limits(max_connections=400, max_keepalive_connections=100))
    return _http


async def recognize_v1(api_key: str, pcm: bytes, lang: str) -> str:
    body = {
        "config": {
            "encoding": "LINEAR16",
            "sampleRateHertz": 16000,
            "languageCode": lang,
            "enableAutomaticPunctuation": True,
            "model": "latest_long",
        },
        "audio": {"content": base64.b64encode(pcm).decode("ascii")},
    }
    # Key in a header, not the URL, so it never lands in access logs.
    r = await http().post(GOOGLE_V1_URL, json=body, headers={"X-Goog-Api-Key": api_key})
    if r.status_code != 200:
        raise RuntimeError(f"google v1 HTTP {r.status_code}")
    results = r.json().get("results", [])
    return " ".join(
        res["alternatives"][0]["transcript"].strip()
        for res in results
        if res.get("alternatives")
    ).strip()


_v2_client = None


async def recognize_v2(cfg: Config, wav: bytes, lang: str) -> str:
    global _v2_client
    from google.api_core.client_options import ClientOptions
    from google.cloud import speech_v2
    from google.cloud.speech_v2.types import cloud_speech

    if _v2_client is None:
        endpoint = None if cfg.location == "global" else f"{cfg.location}-speech.googleapis.com"
        _v2_client = speech_v2.SpeechAsyncClient(
            client_options=ClientOptions(api_endpoint=endpoint) if endpoint else None)
    config = cloud_speech.RecognitionConfig(
        auto_decoding_config=cloud_speech.AutoDetectDecodingConfig(),
        language_codes=[lang],
        model="chirp_2",
        features=cloud_speech.RecognitionFeatures(enable_automatic_punctuation=True),
    )
    req = cloud_speech.RecognizeRequest(
        recognizer=f"projects/{cfg.project}/locations/{cfg.location}/recognizers/_",
        config=config,
        content=wav,
    )
    resp = await _v2_client.recognize(request=req, timeout=30)
    return " ".join(
        r.alternatives[0].transcript.strip() for r in resp.results if r.alternatives
    ).strip()


def create_app(cfg: Config = None, recognizer=None) -> FastAPI:
    """`async recognizer(wav, pcm, lang) -> str` can be injected for tests."""
    cfg = cfg or Config()
    cfg.check()
    limiter = RateLimiter(cfg.rate_per_min)
    in_flight = asyncio.Semaphore(cfg.max_concurrent)

    async def default_recognizer(wav: bytes, pcm: bytes, lang: str) -> str:
        if cfg.backend == "google-v2":
            return await recognize_v2(cfg, wav, lang)
        return await recognize_v1(cfg.api_key, pcm, lang)

    recognize = recognizer or default_recognizer
    app = FastAPI()

    @app.post("/transcribe")
    async def transcribe(
        request: Request,
        x_spell_lang: str = Header(default="nb-NO"),
        x_spell_key: str = Header(default=""),
    ):
        if not hmac.compare_digest(x_spell_key.encode(), cfg.app_key.encode()):
            raise HTTPException(status_code=403, detail="bad app key")
        client_ip = request.client.host if request.client else "?"
        if not limiter.allow(client_ip):
            raise HTTPException(status_code=429, detail="too many requests")
        # Refuse oversized uploads before reading them into memory.
        declared = request.headers.get("content-length")
        if declared and declared.isdigit() and int(declared) > MAX_WAV_BYTES:
            raise HTTPException(status_code=413, detail="audio too long")
        wav = await request.body()
        if len(wav) > MAX_WAV_BYTES:
            raise HTTPException(status_code=413, detail="audio too long")
        try:
            pcm = parse_wav(wav)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        if not pcm:
            return {"text": ""}
        lang = LANG_MAP.get(x_spell_lang.split("-")[0].lower(), x_spell_lang)
        if lang not in ALLOWED_LANGS:
            raise HTTPException(status_code=400, detail="unsupported language")
        if in_flight.locked():
            raise HTTPException(status_code=503, detail="busy, try again")
        try:
            async with in_flight:
                text = await recognize(wav, pcm, lang)
        except Exception:
            # Never echo provider errors (they may contain request details).
            raise HTTPException(status_code=502, detail="speech provider error")
        return {"text": text}

    @app.get("/health")
    async def health():
        return {"ok": True, "backend": cfg.backend}

    return app


# uvicorn stt_proxy:app — tests import the module with SPELL_PROXY_NO_APP=1.
if os.environ.get("SPELL_PROXY_NO_APP") != "1":
    app = create_app()
