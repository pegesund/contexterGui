#!/usr/bin/env python3
"""Spell STT proxy — server-to-server bridge to Google Cloud Speech-to-Text.

Pupil machines POST audio here; THIS server holds the Google credentials.
Clients only know the URL and a rotatable app key — no Google tokens ever
leave the server (see src/stt/cloud.rs for the client side).

Setup:
    pip install fastapi uvicorn google-cloud-speech
    export GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json
    export SPELL_APP_KEY=<choose-a-secret>          # optional but recommended
    uvicorn stt_proxy:app --host 0.0.0.0 --port 8790

Google setup (one-time):
    1. Google Cloud Console → new project → enable "Cloud Speech-to-Text API".
    2. IAM → Service accounts → create → role "Cloud Speech Client".
    3. Keys → add key → JSON → download → point GOOGLE_APPLICATION_CREDENTIALS at it.
    EU data residency: use the eu region endpoint (set below) so pupil audio
    is processed in the EU — relevant for the school DPAs.

Wire format (matches CloudSttEngine):
    POST /transcribe
      Content-Type: audio/wav        (16 kHz mono s16le WAV from the client)
      X-Spell-Lang: nb-NO | nn-NO | en-US  (falls back to nb-NO)
      X-Spell-Key:  <app key>        (checked iff SPELL_APP_KEY is set)
    → 200 {"text": "..."}
"""

import os

from fastapi import FastAPI, Header, HTTPException, Request
from google.api_core.client_options import ClientOptions
from google.cloud import speech_v2
from google.cloud.speech_v2.types import cloud_speech

APP_KEY = os.environ.get("SPELL_APP_KEY", "")
PROJECT = os.environ.get("SPELL_GCP_PROJECT", "")  # set to your project id
# EU processing for GDPR: audio handled in europe-west4.
LOCATION = os.environ.get("SPELL_GCP_LOCATION", "europe-west4")

LANG_MAP = {
    "no": "nb-NO",
    "nb": "nb-NO",
    "nn": "nn-NO",
    "en": "en-US",
}

app = FastAPI()

_client = None


def client() -> speech_v2.SpeechClient:
    global _client
    if _client is None:
        endpoint = f"{LOCATION}-speech.googleapis.com" if LOCATION != "global" else None
        _client = speech_v2.SpeechClient(
            client_options=ClientOptions(api_endpoint=endpoint) if endpoint else None
        )
    return _client


@app.post("/transcribe")
async def transcribe(
    request: Request,
    x_spell_lang: str = Header(default="nb-NO"),
    x_spell_key: str = Header(default=""),
):
    if APP_KEY and x_spell_key != APP_KEY:
        raise HTTPException(status_code=403, detail="bad app key")

    wav = await request.body()
    if len(wav) < 100:
        return {"text": ""}
    # Google STT v2 sync limit is ~1 min audio / 10 MB; our recordings are
    # dictation-length. Larger payloads would need batch/streaming — reject
    # loudly instead of failing opaquely.
    if len(wav) > 10 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="audio too long for sync recognize")

    lang = LANG_MAP.get(x_spell_lang.split("-")[0].lower(), x_spell_lang)

    config = cloud_speech.RecognitionConfig(
        auto_decoding_config=cloud_speech.AutoDetectDecodingConfig(),
        language_codes=[lang],
        model="chirp_2",
        features=cloud_speech.RecognitionFeatures(
            enable_automatic_punctuation=True,
        ),
    )
    req = cloud_speech.RecognizeRequest(
        recognizer=f"projects/{PROJECT}/locations/{LOCATION}/recognizers/_",
        config=config,
        content=wav,
    )
    resp = client().recognize(request=req)
    text = " ".join(
        r.alternatives[0].transcript.strip()
        for r in resp.results
        if r.alternatives
    ).strip()
    return {"text": text}


@app.get("/health")
async def health():
    return {"ok": True}
